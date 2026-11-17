import io
import json
import os
import re
import time
from statistics import median
from typing import Any, Dict, List, Optional, Set, Tuple

import pdfplumber
import requests
import streamlit as st
from docx import Document
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.table import Table
from docx.text.paragraph import Paragraph
from groq import BadRequestError, Groq


# ============================================================
# CONFIGURATION & MRJ TEMPLATE CONSTANTS
# ============================================================

PRIMARY_MODEL = "openai/gpt-oss-120b"
FALLBACK_MODELS = [
    "openai/gpt-oss-20b",
    "llama-3.3-70b-versatile",
]

OPENALEX_URL = "https://api.openalex.org/works"

CURRENT_YEAR = 2026

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

# Flexible section prefix allowing any integer/decimal/Roman numbering
_NUM_PREFIX = r"(?:(?:section\s+)?\d+(?:\.\d+)*\.?|[ivx]+\.?)?\s*"

SECTION_PATTERNS = {
    "Abstract": rf"(?i)^\s*{_NUM_PREFIX}(?:abstract|executive\s+summary)\s*[:.\-—]?$",
    "Introduction": rf"(?i)^\s*{_NUM_PREFIX}(?:introduction|background)\s*[:.\-—]?$",
    "Materials and Methods": (
        rf"(?i)^\s*{_NUM_PREFIX}(?:materials\s+(?:and|&)\s+methods|"
        r"data\s+(?:and|&)\s+methods|methods\s+(?:and|&)\s+materials|"
        r"research\s+methodology|empirical\s+methodology|methodology|"
        r"methods|experimental\s+(?:procedures|setup|design)|"
        r"study\s+area\s+and\s+(?:data|methods|methodology))\s*[:.\-—]?$"
    ),
    "Results and Discussion": (
        rf"(?i)^\s*{_NUM_PREFIX}(?:results\s+(?:and|&)\s+discussion|"
        r"findings\s+(?:and|&)\s+discussion|empirical\s+results(?:\s+and\s+discussion)?|"
        r"analysis\s+and\s+discussion|results|discussion)\s*[:.\-—]?$"
    ),
    "Conclusions": (
        rf"(?i)^\s*{_NUM_PREFIX}(?:conclusions?|summary\s+and\s+conclusions?|"
        r"concluding\s+remarks|conclusion\s+and\s+(?:recommendations|policy\s+implications))\s*[:.\-—]?$"
    ),
    "Multidisciplinary Domains": rf"(?i)^\s*{_NUM_PREFIX}(?:multidisciplinary\s+domains?|research\s+domains?)\s*[:.\-—]?$",
    "Funding": rf"(?i)^\s*{_NUM_PREFIX}(?:funding(?:\s+information)?|financial\s+support|grant\s+support)\s*[:.\-—]?$",
    "Acknowledgments": rf"(?i)^\s*{_NUM_PREFIX}(?:acknowledgments?|acknowledgements?)\s*[:.\-—]?$",
    "Conflicts of Interest": rf"(?i)^\s*{_NUM_PREFIX}(?:conflicts?\s+of\s+interest|competing\s+interests?|disclosure\s+statement)\s*[:.\-—]?$",
    "Declaration on AI Usage": rf"(?i)^\s*{_NUM_PREFIX}(?:declaration\s+(?:on|of)\s+ai(?:\s+usage)?|generative\s+ai\s+statement|ai\s+usage(?:\s+declaration)?|declaration\s+on\s+artificial\s+intelligence)\s*[:.\-—]?$",
    "References": rf"(?i)^\s*{_NUM_PREFIX}(?:references?|bibliography|literature\s+cited)\s*[:.\-—]?$",
    "Ethics Statement": rf"(?i)^\s*{_NUM_PREFIX}(?:ethics\s+statement|ethical\s+(?:approval|statement)|ethics\s+approval|institutional\s+review\s+board|irb\s+statement)\s*[:.\-—]?$",
}

INLINE_SECTION_PREFIXES = {
    "Abstract": r"(?i)^\s*(?:abstract|executive\s+summary)\s*[:\-—]\s+(.+)$",
    "Multidisciplinary Domains": r"(?i)^\s*(?:multidisciplinary\s+domains?|research\s+domains?)\s*[:\-—]\s+(.+)$",
    "Funding": r"(?i)^\s*(?:funding(?:\s+information)?|financial\s+support)\s*[:\-—]\s+(.+)$",
    "Acknowledgments": r"(?i)^\s*(?:acknowledgments?|acknowledgements?)\s*[:\-—]\s+(.+)$",
    "Conflicts of Interest": r"(?i)^\s*(?:conflicts?\s+of\s+interest|competing\s+interests?)\s*[:\-—]\s+(.+)$",
    "Declaration on AI Usage": r"(?i)^\s*(?:declaration\s+(?:on|of)\s+ai(?:\s+usage)?|generative\s+ai\s+statement|ai\s+usage)\s*[:\-—]\s+(.+)$",
    "Ethics Statement": r"(?i)^\s*(?:ethics\s+statement|ethical\s+(?:approval|statement))\s*[:\-—]\s+(.+)$",
}

THESIS_SUBHEADING_PATTERNS = [
    r"(?i)^\s*(?:\d+(?:\.\d+)*\.?\s*)?(?:review of (?:related )?literature|literature review)\s*[:.]?$",
    r"(?i)^\s*(?:\d+(?:\.\d+)*\.?\s*)?(?:statement of (?:the )?problem|problem statement)\s*[:.]?$",
    r"(?i)^\s*(?:\d+(?:\.\d+)*\.?\s*)?(?:hypotheses|hypothesis development)\s*[:.]?$",
    r"(?i)^\s*(?:\d+(?:\.\d+)*\.?\s*)?(?:research questions?)\s*[:.]?$",
    r"(?i)^\s*(?:\d+(?:\.\d+)*\.?\s*)?(?:objectives of (?:the )?study)\s*[:.]?$",
    r"(?i)^\s*(?:\d+(?:\.\d+)*\.?\s*)?(?:delimitations?|scope and delimitations?)\s*[:.]?$",
    r"(?i)^\s*(?:\d+(?:\.\d+)*\.?\s*)?(?:significance of (?:the )?study)\s*[:.]?$",
    r"(?i)^\s*(?:\d+(?:\.\d+)*\.?\s*)?(?:conceptual framework)\s*[:.]?$",
    r"(?i)^\s*(?:\d+(?:\.\d+)*\.?\s*)?(?:operational definitions?)\s*[:.]?$",
]

TEMPLATE_PLACEHOLDER_PATTERNS = [
    r"(?i)\b(?:issn\s*[:\-]?\s*xxxx-xxxx|issn\s+xxxx)\b",
    r"(?i)\b(?:doi\s*[:\-]?\s*10\.xxxx[^\s]*)\b",
    r"(?i)\b(?:volume\s*xx|vol\.\s*xx|issue\s*xx)\b",
    r"(?i)\b(?:received\s*:\s*date|accepted\s*:\s*date|published\s*:\s*date)\b",
    r"(?i)\[insert\s+(?:figure|table|author|text|affiliation)\b.*?\]",
]

PEJORATIVE_OR_IMPRECISE_TERMS = [
    r"\bbackwardness\b",
    r"\bbackward\s+(?:district|districts|region|regions|area|areas|state|states|community|communities|class|classes)\b",
    r"\bprimitive\s+(?:tribe|tribes|society|societies|people|methods)\b",
    r"\buncivilized\b",
]

NON_TAXON_ABBREVS = {
    "e.g", "i.e", "u.s", "u.k", "st", "dr", "mr", "mrs", "ms", "prof",
    "fig", "tab", "vol", "no", "pp", "eq", "et", "vs", "al", "ca", "approx",
}

ASSAM_TERMS = [
    "assam", "guwahati", "gauhati", "tezpur", "silchar", "dibrugarh", "jorhat",
    "diphu", "kokrajhar", "lakhimpur", "nagaon", "barpeta", "bongaigaon", "tinsukia",
    "sivasagar", "sonitpur", "kamrup", "iit guwahati", "indian institute of technology guwahati",
    "gauhati university", "cotton university", "dibrugarh university", "assam university",
    "tezpur university", "nit silchar", "national institute of technology silchar",
    "assam agricultural university", "assam down town university", "assam don bosco",
    "kaziranga university", "bodoland university", "krishna kanta handiqui",
    "csir-neist", "north east institute of science and technology", "iasst",
    "iiit guwahati", "aiims guwahati", "assam engineering college",
]

NORTHEAST_NON_ASSAM_TERMS = [
    "arunachal pradesh", "itanagar", "nerist", "nit arunachal", "rajiv gandhi university",
    "manipur", "imphal", "manipur university", "nit manipur", "cau imphal", "central agricultural university",
    "meghalaya", "shillong", "nehu", "north-eastern hill university", "north eastern hill university",
    "nit meghalaya", "iim shillong", "ri-bhoi",
    "mizoram", "aizawl", "mizoram university", "nit mizoram",
    "nagaland", "kohima", "dimapur", "lumami", "nagaland university", "nit nagaland",
    "sikkim", "gangtok", "tadong", "sikkim university", "nit sikkim", "sikkim manipal",
    "tripura", "agartala", "tripura university", "nit agartala", "mbb university",
]

NORTHEAST_ALL_TERMS = list(set(ASSAM_TERMS + NORTHEAST_NON_ASSAM_TERMS))

GENERIC_INST_TOKENS = {
    "university", "institute", "college", "school", "department", "dept", "centre", "center",
    "research", "technology", "science", "sciences", "engineering", "national", "indian",
    "state", "central", "of", "and", "for", "the", "in", "at", "india",
}

ACADEMIC_STOPWORDS = {
    "a", "an", "the", "and", "or", "but", "in", "on", "at", "to", "for", "with", "of",
    "by", "from", "using", "use", "used", "study", "studies", "analysis", "analyze",
    "analysed", "analyzed", "approach", "approaches", "based", "case", "india", "indian",
    "paper", "research", "results", "result", "method", "methods", "methodology",
    "data", "dataset", "datasets", "model", "models", "proposed", "propose", "present",
    "presents", "presented", "show", "shows", "showed", "shown", "finding", "findings",
    "conclusion", "conclusions", "significant", "significantly", "observed", "observe",
    "between", "among", "across", "within", "through", "into", "over", "under", "high",
    "low", "higher", "lower", "increase", "increased", "decrease", "decreased", "different",
    "various", "several", "two", "three", "four", "five", "also", "however", "therefore",
    "thus", "moreover", "furthermore", "whereas", "while", "during", "well", "new",
    "novel", "effective", "efficient", "performance", "evaluation", "evaluate", "evaluated",
    "compared", "comparison", "investigate", "investigated", "examine", "examined",
    "determine", "determined", "assess", "assessed", "assessment", "impact", "effect",
    "effects", "role", "factors", "factor", "level", "levels", "rate", "rates", "area",
    "areas", "region", "regions", "regional", "state", "district", "districts", "sample",
    "samples", "values", "value", "table", "figure", "abstract", "keywords",
}


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
    clean = re.sub(r"^```(?:json)?\s*", "", (raw_resp or "").strip(), flags=re.MULTILINE)
    clean = re.sub(r"```\s*$", "", clean.strip(), flags=re.MULTILINE).strip()
    try:
        return json.loads(clean)
    except Exception:
        match = re.search(r"(\{.*\})", clean, re.DOTALL)
        if match:
            return json.loads(match.group(1))
        raise ValueError("Could not extract valid JSON from LLM audit response.")


def chunk_text(text: str, max_chars: int = 10000, overlap: int = 800) -> List[str]:
    text = (text or "").strip()
    if not text:
        return []
    if len(text) <= max_chars:
        return [text]

    chunks = []
    start = 0
    n = len(text)
    while start < n:
        end = min(start + max_chars, n)
        if end < n:
            split_pos = text.rfind("\n", start + int(max_chars * 0.7), end)
            if split_pos != -1:
                end = split_pos
        chunks.append(text[start:end].strip())
        if end >= n:
            break
        start = max(end - overlap, start + 1)
    return chunks


# ============================================================
# DOCUMENT EXTRACTION (.DOCX READING ORDER + .PDF)
# ============================================================

def _extract_docx_textbox_texts(element) -> List[str]:
    texts = []
    for txbx in element.findall(".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}txbxContent"):
        p_texts = []
        for p in txbx.findall(".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p"):
            t_parts = [t.text for t in p.findall(".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t") if t.text]
            joined = normalize("".join(t_parts))
            if joined:
                p_texts.append(joined)
        if p_texts:
            texts.extend(p_texts)
    return texts


def _scan_unitalicized_binomials_in_para(para: Paragraph) -> List[str]:
    flags = []
    full_txt = para.text or ""
    if not full_txt.strip():
        return flags

    taxon_pattern = re.compile(r"\b([A-Z]\.\s+[a-z]{4,})\b")
    matches = list(taxon_pattern.finditer(full_txt))
    if not matches:
        return flags

    italic_map = []
    for run in para.runs:
        r_txt = run.text or ""
        is_it = bool(run.italic)
        italic_map.extend([is_it] * len(r_txt))

    for m in matches:
        phrase = m.group(1)
        prefix = phrase.split(".")[0].lower()
        if prefix in NON_TAXON_ABBREVS:
            continue
        start, end = m.span(1)
        if end <= len(italic_map):
            span_italics = italic_map[start:end]
            if span_italics and not all(span_italics):
                flags.append(f"Unitalicized taxonomic name in DOCX: '{phrase}'")
    return flags


def extract_text_and_assets(uploaded_file) -> Tuple[str, Dict[str, Any]]:
    data = uploaded_file.getvalue()
    name = uploaded_file.name.lower()
    asset_meta: Dict[str, Any] = {
        "file_type": "docx" if name.endswith(".docx") else "pdf",
        "total_images": 0,
        "total_tables": 0,
        "detected_captions": [],
        "in_text_figure_refs": [],
        "in_text_table_refs": [],
        "table_inventory": [],
        "unitalicized_taxa_xml": [],
    }

    seen_labels = set()
    caption_regex = re.compile(r"^\s*((?:Figure|Fig\.?|Table)\s*\d+[A-Za-z]?)\s*[:.\-—\s]\s*(.+)$", re.IGNORECASE)

    def register_caption(line_txt: str, has_visual_hint: bool):
        m = caption_regex.match(line_txt)
        if not m:
            return
        raw_label = normalize(m.group(1)).title().replace("Fig.", "Figure").replace("Fig ", "Figure ")
        if raw_label not in seen_labels:
            seen_labels.add(raw_label)
            asset_meta["detected_captions"].append({
                "label": raw_label,
                "caption": line_txt[:220],
                "placement": "In-line document body",
                "has_visual": has_visual_hint,
            })

    if name.endswith(".docx"):
        doc = Document(io.BytesIO(data))
        parts: List[str] = []
        asset_meta["total_tables"] = len(doc.tables)
        xml_text = doc._element.xml
        asset_meta["total_images"] = len(re.findall(r"<a:blip|<w:drawing", xml_text))

        table_idx = 0
        for child in doc.element.body.iterchildren():
            if isinstance(child, CT_P):
                p = Paragraph(child, doc)
                txt = p.text.strip()
                if txt:
                    parts.append(txt)
                    is_fig = bool(re.match(r"^\s*(?:Fig|Figure)", txt, re.I))
                    register_caption(
                        txt,
                        has_visual_hint=(asset_meta["total_images"] > 0 if is_fig else asset_meta["total_tables"] > 0),
                    )
                asset_meta["unitalicized_taxa_xml"].extend(_scan_unitalicized_binomials_in_para(p))

                for tb_line in _extract_docx_textbox_texts(child):
                    parts.append(tb_line)
                    is_fig = bool(re.match(r"^\s*(?:Fig|Figure)", tb_line, re.I))
                    register_caption(
                        tb_line,
                        has_visual_hint=(asset_meta["total_images"] > 0 if is_fig else asset_meta["total_tables"] > 0),
                    )

            elif isinstance(child, CT_Tbl):
                table = Table(child, doc)
                table_idx += 1
                rows_text = []
                for r_i, row in enumerate(table.rows):
                    row_cells = [normalize(cell.text) for cell in row.cells]
                    row_joined = " | ".join(row_cells)
                    if any(row_cells):
                        rows_text.append(row_joined)
                        if r_i == 0 and caption_regex.match(row_cells[0]):
                            register_caption(row_cells[0], has_visual_hint=True)
                if rows_text:
                    header_preview = rows_text[0][:140]
                    asset_meta["table_inventory"].append({
                        "table_index": table_idx,
                        "rows": len(table.rows),
                        "cols": len(table.columns),
                        "header_preview": header_preview,
                    })
                    parts.append(f"[TABLE {table_idx} | {len(table.rows)}x{len(table.columns)}]\n" + "\n".join(rows_text))

        full_text = "\n".join(parts)

    elif name.endswith(".pdf"):
        pages = []
        with pdfplumber.open(io.BytesIO(data)) as pdf:
            for page in pdf.pages:
                page_text = page.extract_text() or ""
                pages.append(page_text)
                images_on_page = len(page.images)
                tables_on_page = page.extract_tables() or []
                asset_meta["total_images"] += images_on_page
                asset_meta["total_tables"] += len(tables_on_page)

                for t in tables_on_page:
                    if t and t[0]:
                        asset_meta["table_inventory"].append({
                            "table_index": len(asset_meta["table_inventory"]) + 1,
                            "rows": len(t),
                            "cols": len(t[0]),
                            "header_preview": normalize(" | ".join([str(c or "") for c in t[0]]))[:140],
                        })

                for line in page_text.splitlines():
                    sline = line.strip()
                    if sline:
                        is_fig = bool(re.match(r"^\s*(?:Fig|Figure)", sline, re.I))
                        register_caption(
                            sline,
                            has_visual_hint=(
                                (images_on_page > 0 or asset_meta["total_images"] > 0)
                                if is_fig
                                else (len(tables_on_page) > 0 or asset_meta["total_tables"] > 0)
                            ),
                        )
        full_text = "\n".join(pages)
    else:
        raise ValueError("Unsupported format. Please upload a .docx or .pdf file.")

    fig_refs = sorted(set(re.findall(r"\b(?:Figure|Fig\.?)\s+(\d+[A-Za-z]?)\b", full_text, re.I)))
    tab_refs = sorted(set(re.findall(r"\bTable\s+(\d+[A-Za-z]?)\b", full_text, re.I)))
    asset_meta["in_text_figure_refs"] = [f"Figure {x}" for x in fig_refs]
    asset_meta["in_text_table_refs"] = [f"Table {x}" for x in tab_refs]
    asset_meta["unitalicized_taxa_xml"] = sorted(set(asset_meta["unitalicized_taxa_xml"]))

    return full_text, asset_meta


# ============================================================
# DETERMINISTIC SECTION SPLITTER, AUTHOR EXTRACTOR & PRE-SCANNER
# ============================================================

def _match_section_heading(line: str) -> Optional[str]:
    clean = normalize(line)
    if not clean:
        return None
    if len(clean) <= 95 and word_count(clean) <= 10:
        for sec_name, pattern in SECTION_PATTERNS.items():
            if re.match(pattern, clean):
                return sec_name
    for sec_name, pattern in INLINE_SECTION_PREFIXES.items():
        if re.match(pattern, clean):
            return sec_name
    return None


def split_manuscript_into_sections(text: str) -> Tuple[Dict[str, str], Dict[str, Dict[str, Any]]]:
    lines = text.splitlines()
    detected_map: Dict[str, Dict[str, Any]] = {}
    ordered_hits: List[Tuple[int, str, str]] = []

    for i, line in enumerate(lines):
        clean = normalize(line)
        sec_name = _match_section_heading(clean)
        if sec_name and sec_name not in detected_map:
            quote = ""
            if sec_name in INLINE_SECTION_PREFIXES:
                m_in = re.match(INLINE_SECTION_PREFIXES[sec_name], clean)
                if m_in:
                    quote = normalize(m_in.group(1))[:160]
            if not quote:
                for nxt in lines[i + 1: i + 6]:
                    cleaned_nxt = normalize(nxt)
                    if cleaned_nxt and not _match_section_heading(cleaned_nxt):
                        quote = cleaned_nxt[:160]
                        break
            detected_map[sec_name] = {
                "heading": clean[:95],
                "first_line_quote": quote or "Section heading identified.",
                "line_num": i,
            }
            ordered_hits.append((i, sec_name, clean))

    if "Abstract" not in detected_map:
        intro_line = detected_map.get("Introduction", {}).get("line_num", len(lines))
        search_limit = min(len(lines), intro_line, 55)
        for idx in range(search_limit):
            l = normalize(lines[idx])
            if (
                word_count(l) >= 35
                and not re.search(r"(?i)@|department\b|university\b|institute\b|received:|accepted:|doi\.org|issn|vol\.\s*\d+", l)
                and not _match_section_heading(l)
            ):
                detected_map["Abstract"] = {
                    "heading": "Implicit Abstract (Heading omitted)",
                    "first_line_quote": l[:160],
                    "line_num": idx,
                }
                ordered_hits.append((idx, "Abstract", "Implicit Abstract"))
                ordered_hits.sort(key=lambda x: x[0])
                break

    sections_text: Dict[str, str] = {}
    if ordered_hits:
        first_idx = ordered_hits[0][0]
        sections_text["FrontMatter"] = "\n".join(lines[:first_idx]).strip()
        for k, (line_idx, sec_name, _) in enumerate(ordered_hits):
            end_idx = ordered_hits[k + 1][0] if k + 1 < len(ordered_hits) else len(lines)
            sections_text[sec_name] = "\n".join(lines[line_idx:end_idx]).strip()
    else:
        sections_text["FrontMatter"] = ""
        sections_text["FullBody"] = text

    return sections_text, detected_map


def extract_author_metadata_deterministic(front_matter: str, full_text: str) -> Dict[str, List[str]]:
    lines = [normalize(l) for l in front_matter.splitlines() if normalize(l)]
    affiliations: List[str] = []
    authors: List[str] = []

    inst_regex = re.compile(
        r"(?i)\b(university|institute|college|iit|nit|iiit|aiims|school of|department of|faculty of|academy|laboratory|centre for|center for)\b"
    )
    meta_noise_regex = re.compile(
        r"(?i)^(received|accepted|published|issn|doi|keywords|correspondence|running title|abstract|original\s+(?:research\s+)?article|research\s+paper)"
    )

    for idx, line in enumerate(lines):
        if meta_noise_regex.match(line):
            continue
        if inst_regex.search(line):
            chunks = re.split(r"\s*;\s*|\s*\d+\s*(?=[A-Z])", line)
            for ch in chunks:
                ch_clean = normalize(ch).strip(" ,.;:*†1234567890")
                if len(ch_clean) >= 6 and inst_regex.search(ch_clean):
                    affiliations.append(ch_clean)
        elif idx > 0 and word_count(line) <= 22 and not re.search(r"@|\d{4}|http", line):
            if "," in line or " and " in line.lower() or idx <= 3:
                cleaned_line = re.sub(r"[*†‡§\d]+", "", line)
                for cand in re.split(r"\s*,\s*|\s+and\s+", cleaned_line, flags=re.I):
                    c = normalize(cand)
                    if 2 <= word_count(c) <= 5 and not inst_regex.search(c):
                        authors.append(c)

    for line in full_text.splitlines()[:60]:
        if re.search(r"(?i)corresponding author|affiliation", line) and inst_regex.search(line):
            affiliations.append(normalize(line))

    return {
        "authors": list(dict.fromkeys(authors)),
        "affiliations": list(dict.fromkeys(affiliations)),
    }


def analyze_references_currency(ref_text: str) -> Dict[str, Any]:
    if not ref_text:
        return {
            "total_references_estimated": 0,
            "years_detected_count": 0,
            "median_year": None,
            "oldest_year": None,
            "newest_year": None,
            "pct_older_than_10_yrs": 0.0,
            "pct_older_than_15_yrs": 0.0,
            "pct_recent_5_yrs": 0.0,
        }

    lines = [l.strip() for l in ref_text.splitlines() if len(l.strip()) > 20]
    years = [int(y) for y in re.findall(r"\b(19\d{2}|20[0-2]\d)\b", ref_text) if 1940 <= int(y) <= CURRENT_YEAR]
    if not years:
        return {
            "total_references_estimated": max(len(lines) - 1, 0),
            "years_detected_count": 0,
            "median_year": None,
            "oldest_year": None,
            "newest_year": None,
            "pct_older_than_10_yrs": 0.0,
            "pct_older_than_15_yrs": 0.0,
            "pct_recent_5_yrs": 0.0,
        }

    older_10 = sum(1 for y in years if (CURRENT_YEAR - y) > 10)
    older_15 = sum(1 for y in years if (CURRENT_YEAR - y) > 15)
    recent_5 = sum(1 for y in years if (CURRENT_YEAR - y) <= 5)
    n = len(years)
    return {
        "total_references_estimated": max(len(lines) - 1, n),
        "years_detected_count": n,
        "median_year": int(median(years)),
        "oldest_year": min(years),
        "newest_year": max(years),
        "pct_older_than_10_yrs": round(100.0 * older_10 / n, 1),
        "pct_older_than_15_yrs": round(100.0 * older_15 / n, 1),
        "pct_recent_5_yrs": round(100.0 * recent_5 / n, 1),
    }


def scan_scholarly_and_statistical_signals(text: str, asset_meta: Dict[str, Any]) -> Dict[str, Any]:
    signals: Dict[str, Any] = {
        "improper_p_values_detected": [],
        "pearson_directional_fallacy_hits": [],
        "pseudoreplication_context_hits": [],
        "pejorative_terms_detected": [],
        "old_dataset_years_detected": [],
        "typo_concatenation_hits": [],
        "unitalicized_taxa_xml": asset_meta.get("unitalicized_taxa_xml", []),
    }

    p_zero = re.findall(r"(?i)\bp(?:\s*-?\s*value)?\s*(?:=|is)\s*0?(?:\.0+)?(?!\d)", text)
    if p_zero:
        signals["improper_p_values_detected"] = sorted(set(normalize(m) for m in p_zero))

    sentences = re.split(r"(?<=[.!?])\s+|\n+", text)
    for idx, sent in enumerate(sentences):
        window = " ".join(sentences[max(0, idx - 1): min(len(sentences), idx + 2)])
        if re.search(r"(?i)\b(?:pearson|spearman|correlation)\b", window) and re.search(
            r"(?i)\b(?:dependent\s+variable|independent\s+variables?)\b", sent
        ):
            signals["pearson_directional_fallacy_hits"].append(normalize(sent)[:200])

        if re.search(r"(?i)\b(?:pixels?|voxels?|technical\s+replicates?|sub-?samples?)\b", window) and re.search(
            r"(?i)\b(?:anova|t-?test|chi-?square|χ²|p\s*[<>=]|degrees?\s+of\s+freedom)\b", sent
        ):
            signals["pseudoreplication_context_hits"].append(normalize(sent)[:200])

    signals["pearson_directional_fallacy_hits"] = signals["pearson_directional_fallacy_hits"][:5]
    signals["pseudoreplication_context_hits"] = signals["pseudoreplication_context_hits"][:5]

    for pat in PEJORATIVE_OR_IMPRECISE_TERMS:
        for m in re.findall(pat, text, flags=re.I):
            signals["pejorative_terms_detected"].append(normalize(m))
    signals["pejorative_terms_detected"] = sorted(set(signals["pejorative_terms_detected"]))

    old_ds = re.findall(r"(?i)\b(?:census\s+(?:of\s+)?(19\d{2}|200\d|201[0-5])|(19\d{2}|200\d|201[0-5])\s+census)\b", text)
    flat_ds = [y for tup in old_ds for y in tup if y]
    if flat_ds:
        signals["old_dataset_years_detected"] = sorted(set(flat_ds))

    concat_hits = re.findall(r"\b[a-z]{4,}[A-Z]{2,}[A-Za-z]*\b", text)
    spell_hits = re.findall(r"(?i)\b(?:dimentional|reasearch|signficant|enviroment|analaysis|methdology)\b", text)
    signals["typo_concatenation_hits"] = sorted(set(concat_hits[:8] + spell_hits[:8]))

    return signals


def preaudit_mrj_template(
    text: str,
    asset_meta: Dict[str, Any],
) -> Tuple[List[Dict[str, Any]], List[str], List[str], Dict[str, Any], Dict[str, str]]:
    sections_text, detected_map = split_manuscript_into_sections(text)
    lines = text.splitlines()

    detected_thesis_headers: List[str] = []
    detected_placeholders: List[str] = []

    for ph_pattern in TEMPLATE_PLACEHOLDER_PATTERNS:
        for m in re.findall(ph_pattern, text):
            clean_m = normalize(m)
            if clean_m and clean_m not in detected_placeholders:
                detected_placeholders.append(clean_m)

    for line in lines:
        clean = normalize(line)
        if not clean or len(clean) > 95:
            continue
        for t_pattern in THESIS_SUBHEADING_PATTERNS:
            if re.match(t_pattern, clean) and clean not in detected_thesis_headers:
                detected_thesis_headers.append(clean)

    abstract_raw = sections_text.get("Abstract", "")
    if abstract_raw:
        abs_clean = re.sub(r"(?is)^\s*(?:abstract|executive\s+summary)\s*[:.\-—]?\s*", "", abstract_raw)
        abs_clean = re.split(r"(?im)^\s*keywords?\s*[:\-—]", abs_clean)[0].strip()
        abstract_wc = word_count(abs_clean)
    else:
        abs_clean = ""
        abstract_wc = 0

    keywords_detected: List[str] = []
    m_kw = re.search(
        r"(?is)\bkeywords?\s*[:\-—]\s*(.*?)(?=\n\s*(?:(?:\d+\.?\s+)?introduction|materials|data\s+and\s+methods|abstract)\b|$)",
        text,
    )
    if m_kw:
        raw_kw = m_kw.group(1).strip().splitlines()[0]
        keywords_detected = [normalize(k) for k in re.split(r"[;,]", raw_kw) if len(normalize(k)) > 1]

    dom_found = False
    dom_items: List[str] = []
    dom_section = sections_text.get("Multidisciplinary Domains", "")
    if not dom_section:
        m_dom = re.search(
            r"(?is)\bmultidisciplinary\s+domains?\b.*?(?=\n\s*(?:funding|acknowledg|conflicts|declaration|references|ethics)\b|$)",
            text,
        )
        dom_section = m_dom.group(0) if m_dom else ""

    if dom_section:
        dom_found = True
        raw_items = re.findall(
            r"\([a-z0-9]\)\s*([^(\n.;]+?)(?=\s*\([a-z0-9]\)|[.;\n]|$)",
            dom_section,
            flags=re.I,
        )
        for item in raw_items:
            cleaned_item = re.sub(r"(?i)\s+(?:and|or)$", "", normalize(item)).strip(" ,")
            if len(cleaned_item) >= 3:
                dom_items.append(cleaned_item)

    ref_stats = analyze_references_currency(sections_text.get("References", ""))
    scholarly_signals = scan_scholarly_and_statistical_signals(text, asset_meta)
    author_meta = extract_author_metadata_deterministic(sections_text.get("FrontMatter", ""), text)

    template_observations: Dict[str, Any] = {
        "abstract_text": abs_clean,
        "abstract_word_count": abstract_wc,
        "keywords_detected": keywords_detected,
        "multidisciplinary_domains_found": dom_found,
        "multidisciplinary_domains_list": dom_items,
        "multidisciplinary_domains_count": len(dom_items),
        "declaration_ai_usage_found": "Declaration on AI Usage" in detected_map
        or bool(re.search(r"(?is)\bdeclaration\s+(?:on|of)\s+ai\s+usage\b|prepared\s+without\s+the\s+use\s+of\s+ai\s+tools", text)),
        "funding_statement_found": "Funding" in detected_map,
        "conflicts_statement_found": "Conflicts of Interest" in detected_map,
        "square_bracket_citations_found": bool(re.findall(r"\[\d+(?:[–\-–,]\s*\d+)*\]", text)),
        "reference_currency": ref_stats,
        "scholarly_signals": scholarly_signals,
        "author_metadata": author_meta,
        "total_characters": len(text),
        "total_words": word_count(text),
    }

    preaudited: List[Dict[str, Any]] = []
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

    return preaudited, detected_thesis_headers, detected_placeholders, template_observations, sections_text


# ============================================================
# STRICT AUDIT JSON SCHEMA SPECIFICATION
# ============================================================

AUDIT_STRICT_SCHEMA = {
    "name": "manuscript_editorial_audit",
    "strict": True,
    "schema": {
        "type": "object",
        "properties": {
            "manuscript_title": {"type": "string"},
            "extracted_authors": {"type": "array", "items": {"type": "string"}},
            "extracted_affiliations": {"type": "array", "items": {"type": "string"}},
            "reviewer_search_queries": {"type": "array", "items": {"type": "string"}},
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
                    "abstract_claimed_items": {"type": "array", "items": {"type": "string"}},
                    "verified_in_body_items": {"type": "array", "items": {"type": "string"}},
                    "abstract_body_discrepancies": {"type": "array", "items": {"type": "string"}},
                },
                "required": [
                    "abstract_claimed_items",
                    "verified_in_body_items",
                    "abstract_body_discrepancies",
                ],
                "additionalProperties": False,
            },
            "pass_2_template_and_layout": {
                "type": "object",
                "properties": {
                    "journal_style_compliance": {"type": "string", "enum": ["PASS", "FAIL"]},
                    "unwanted_thesis_subheadings": {"type": "array", "items": {"type": "string"}},
                    "template_placeholders_detected": {"type": "array", "items": {"type": "string"}},
                    "multidisciplinary_domains_compliance": {"type": "string", "enum": ["PASS", "FAIL"]},
                    "abstract_quality_and_word_count": {"type": "string"},
                    "reference_currency_and_table_originality": {"type": "array", "items": {"type": "string"}},
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
                    "reference_currency_and_table_originality",
                    "structural_section_checks",
                ],
                "additionalProperties": False,
            },
            "pass_3_statistical_and_mathematical_rigor": {
                "type": "object",
                "properties": {
                    "p_value_reporting_issues": {"type": "array", "items": {"type": "string"}},
                    "pseudoreplication_flags": {"type": "array", "items": {"type": "string"}},
                    "statistical_fallacies_and_test_validity": {"type": "array", "items": {"type": "string"}},
                    "statistical_test_completeness": {"type": "string"},
                },
                "required": [
                    "p_value_reporting_issues",
                    "pseudoreplication_flags",
                    "statistical_fallacies_and_test_validity",
                    "statistical_test_completeness",
                ],
                "additionalProperties": False,
            },
            "pass_4_domain_and_reproducibility": {
                "type": "object",
                "properties": {
                    "data_vintage_and_novelty_notes": {"type": "array", "items": {"type": "string"}},
                    "terminology_and_proxy_validity_flags": {"type": "array", "items": {"type": "string"}},
                    "conceptual_overinterpretation_flags": {"type": "array", "items": {"type": "string"}},
                    "is_computational_or_algorithm_paper": {"type": "boolean"},
                    "software_novelty_and_mapping_notes": {"type": "array", "items": {"type": "string"}},
                    "algorithmic_edge_cases_notes": {"type": "array", "items": {"type": "string"}},
                    "software_dependency_versions": {"type": "string"},
                    "benchmarks_and_scalability": {"type": "string"},
                },
                "required": [
                    "data_vintage_and_novelty_notes",
                    "terminology_and_proxy_validity_flags",
                    "conceptual_overinterpretation_flags",
                    "is_computational_or_algorithm_paper",
                    "software_novelty_and_mapping_notes",
                    "algorithmic_edge_cases_notes",
                    "software_dependency_versions",
                    "benchmarks_and_scalability",
                ],
                "additionalProperties": False,
            },
            "pass_5_literal_text_and_captions": {
                "type": "object",
                "properties": {
                    "enumeration_figure_inconsistencies": {"type": "array", "items": {"type": "string"}},
                    "taxonomic_formatting_flags": {"type": "array", "items": {"type": "string"}},
                    "typographical_and_spacing_issues": {"type": "array", "items": {"type": "string"}},
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
            "extracted_authors",
            "extracted_affiliations",
            "reviewer_search_queries",
            "editorial_verdict",
            "pass_1_cross_section_integrity",
            "pass_2_template_and_layout",
            "pass_3_statistical_and_mathematical_rigor",
            "pass_4_domain_and_reproducibility",
            "pass_5_literal_text_and_captions",
        ],
        "additionalProperties": False,
    },
}


# ============================================================
# RESILIENT GROQ API CLIENT (CHUNKED MAP-REDUCE, ZERO TRUNCATION)
# ============================================================

def call_groq_json(
    client: Groq,
    messages: List[Dict[str, str]],
    schema: Optional[Dict[str, Any]] = None,
    max_tokens: int = 3000,
    primary_model: str = PRIMARY_MODEL,
) -> Dict[str, Any]:
    models_to_try = [primary_model] + [m for m in FALLBACK_MODELS if m != primary_model]

    for model_name in models_to_try:
        is_gpt_oss = "gpt-oss" in model_name
        attempts = 3
        for attempt in range(attempts):
            try:
                kwargs: Dict[str, Any] = {
                    "model": model_name,
                    "messages": messages,
                    "temperature": 0.0,
                    "max_completion_tokens": max_tokens,
                }
                if is_gpt_oss:
                    kwargs["reasoning_effort"] = "low"
                    kwargs["include_reasoning"] = False

                if schema and is_gpt_oss and attempt == 0:
                    kwargs["response_format"] = {"type": "json_schema", "json_schema": schema}
                else:
                    kwargs["response_format"] = {"type": "json_object"}

                resp = client.chat.completions.create(**kwargs)
                return clean_json_response(resp.choices[0].message.content or "{}")

            except BadRequestError:
                if attempt == 0 and schema:
                    continue
                break
            except Exception as err:
                err_str = str(err).lower()
                if "429" in err_str or "rate" in err_str or "tokens per minute" in err_str:
                    wait_sec = 12 * (attempt + 1)
                    time.sleep(wait_sec)
                    continue
                if "413" in err_str or "too large" in err_str:
                    break
                if attempt == attempts - 1:
                    break
                time.sleep(2)

    raise RuntimeError("All Groq model attempts failed. Verify API key or reduce chunk size.")


def run_editorial_audit(
    raw_text: str,
    asset_meta: Dict[str, Any],
    client: Groq,
    free_tier_mode: bool = True,
    progress_cb=None,
) -> Tuple[Dict[str, Any], List[Dict[str, Any]], Dict[str, Any]]:
    preaudited, detected_thesis, detected_placeholders, template_obs, sections_text = preaudit_mrj_template(
        raw_text, asset_meta
    )

    chunk_char_limit = 9500 if free_tier_mode else 45000

    front_intro = "\n\n".join(
        s for s in [
            sections_text.get("FrontMatter", ""),
            sections_text.get("Abstract", ""),
            sections_text.get("Introduction", ""),
            sections_text.get("Multidisciplinary Domains", ""),
            sections_text.get("Funding", ""),
            sections_text.get("Acknowledgments", ""),
            sections_text.get("Conflicts of Interest", ""),
            sections_text.get("Declaration on AI Usage", ""),
            sections_text.get("Ethics Statement", ""),
        ] if s
    ).strip()

    body_methods_results = "\n\n".join(
        s for s in [
            sections_text.get("Materials and Methods", ""),
            sections_text.get("Results and Discussion", ""),
            sections_text.get("Conclusions", ""),
        ] if s
    ).strip()

    if not body_methods_results and not front_intro:
        body_methods_results = raw_text
        front_intro = raw_text[:6000]
    elif not body_methods_results:
        body_methods_results = raw_text

    ground_truth_dossier = {
        "total_words": template_obs["total_words"],
        "detected_sections": preaudited,
        "unwanted_thesis_subheadings": detected_thesis,
        "placeholders_detected": detected_placeholders,
        "abstract_word_count": template_obs["abstract_word_count"],
        "multidisciplinary_domains_found": template_obs["multidisciplinary_domains_found"],
        "multidisciplinary_domains_list": template_obs["multidisciplinary_domains_list"],
        "reference_currency": template_obs["reference_currency"],
        "deterministic_scholarly_signals": template_obs["scholarly_signals"],
        "visual_inventory": {
            "total_images": asset_meta["total_images"],
            "total_tables": asset_meta["total_tables"],
            "detected_captions": asset_meta["detected_captions"],
            "in_text_figure_refs": asset_meta["in_text_figure_refs"],
            "in_text_table_refs": asset_meta["in_text_table_refs"],
            "table_inventory": asset_meta["table_inventory"],
        },
    }

    if progress_cb:
        progress_cb("Stage 1/3: Auditing Front Matter, Abstract Claims, Authors & Conceptual Framing...")

    front_chunks = chunk_text(front_intro, max_chars=chunk_char_limit, overlap=500) or [raw_text[:chunk_char_limit]]
    stage1_notes = []
    for idx, f_chunk in enumerate(front_chunks):
        s1_prompt = f"""Analyze this Front Matter / Abstract / Introduction / Declarations excerpt (Chunk {idx+1}/{len(front_chunks)}) for an academic manuscript audit.
Return valid JSON with keys:
- "manuscript_title": exact title of the manuscript
- "authors": list of ALL author full names found
- "affiliations": list of ALL author institutions/universities/departments found
- "reviewer_search_queries": 4 to 6 crisp, domain-specific academic search queries (2 to 4 words each, NO punctuation, NO geographical places unless central to the topic) derived STRICTLY from the Manuscript Title and Abstract so we can locate peer reviewers working on this exact topic
- "abstract_claimed_items": list of EVERY dataset, experiment, algorithm, method, variable, or organism claimed in the Abstract
- "data_vintage_and_novelty_notes": check dataset years (e.g. 2011 Census analyzed in {CURRENT_YEAR}) and whether novelty is justified
- "terminology_and_proxy_validity_flags": check for pejorative/imprecise terms (e.g. 'backwardness') and weak proxy variables (e.g. sex ratio alone as a primary development proxy instead of HDI/MPI/GII/IMR/safe drinking water)
- "citation_currency_notes": note outdated literature in Introduction using reference stats: {json.dumps(template_obs['reference_currency'])}
- "text_and_typo_flags": list of typos or formatting issues seen here

EXCERPT:
{f_chunk}"""
        s1_res = call_groq_json(
            client,
            [
                {"role": "system", "content": "You are an expert academic peer-review auditor. Return strictly valid JSON."},
                {"role": "user", "content": s1_prompt},
            ],
            max_tokens=1900,
        )
        stage1_notes.append(s1_res)
        if free_tier_mode and idx < len(front_chunks) - 1:
            time.sleep(4)

    claimed_items = []
    for n in stage1_notes:
        claimed_items.extend(n.get("abstract_claimed_items", []))
    claimed_items = list(dict.fromkeys(claimed_items))

    body_chunks = chunk_text(body_methods_results, max_chars=chunk_char_limit, overlap=800)
    stage2_notes = []
    for idx, b_chunk in enumerate(body_chunks):
        if progress_cb:
            progress_cb(f"Stage 2/3: Deep-auditing Methods, Results & Statistics (Chunk {idx+1}/{len(body_chunks)})...")
        if free_tier_mode:
            time.sleep(5)

        s2_prompt = f"""Audit this Methods / Results / Discussion / Conclusions chunk (Chunk {idx+1} of {len(body_chunks)}) thoroughly.
ABSTRACT CLAIMS TO VERIFY IN THIS CHUNK:
{json.dumps(claimed_items, indent=2)}

Check strictly for ALL of the following:
1. "verified_claims_found_in_chunk": which Abstract claims have explicit methodology, data, or empirical results in THIS chunk?
2. "unverified_or_suspicious_claims": any contradictions or missing details noted in this chunk.
3. "p_value_issues": improper reporting (e.g., p = 0.00 or p = 0 instead of p < 0.001).
4. "pseudoreplication_flags": are pixels, image gradients, voxels, technical replicates, or subsamples treated as independent observational units (n)? Is the observational unit ambiguous?
5. "statistical_fallacies_and_test_validity":
   - Symmetric vs. directional fallacy: e.g., calling variables "dependent" and "independent" in Pearson/Spearman correlation!
   - Chi-square / ANOVA validity: missing contingency table construction, observed/expected frequency justification, degrees of freedom rationale, or assumption checks.
6. "table_originality_flags": do tables merely replicate raw secondary/Census tables without analytical value?
7. "conceptual_overinterpretation_flags": equating image-processing/computational metrics (e.g. image gradient intensity) directly to biological complexity or evolutionary constraints; unsubstantiated O(n^2) complexity, mathematical uniqueness, or alignment-free superiority claims.
8. "is_computational_or_algorithm_paper": boolean.
9. "software_novelty_and_mapping_notes": for software/tool papers, is novelty explicitly differentiated from existing tools? Are exact mathematical coordinate/RGB encoding formulas and sequences < k documented?
10. "algorithmic_edge_cases_notes": ambiguous bases ('N'), case sensitivity, overlapping vs non-overlapping k-mers, reverse complement, multi-chromosomal handling, dependency versions, runtime/memory benchmarks.
11. "literal_text_and_caption_flags": enumeration vs figure mismatches, unitalicized taxonomic binomials, typos/concatenated words.

Return strictly valid JSON with those exact keys.

CHUNK {idx+1}/{len(body_chunks)}:
{b_chunk}"""

        s2_res = call_groq_json(
            client,
            [
                {"role": "system", "content": "You are a rigorous statistical, methodological, and domain peer reviewer. Return strictly valid JSON."},
                {"role": "user", "content": s2_prompt},
            ],
            max_tokens=2200,
        )
        stage2_notes.append(s2_res)

    if progress_cb:
        progress_cb("Stage 3/3: Synthesizing 5-Pass Audit & Reconciling Ground-Truth Evidence...")
    if free_tier_mode:
        time.sleep(5)

    reduce_prompt = f"""You are the Lead Academic Editor for Multidisciplinary Research Journal (MRJ).
Synthesize the Deterministic Ground-Truth Dossier, Stage 1 Front-Matter Audit, and Stage 2 Full-Body Chunks into the final strict JSON audit report.

CRITICAL GROUND-TRUTH RULES (DO NOT VIOLATE):
1. NEVER mark a section as FAIL or 'not identified' if the Deterministic Dossier shows status_hint == 'PASS'.
2. NEVER claim tables or figures are missing if the Visual Inventory shows them present.
3. NEVER mark 'multidisciplinary_domains_compliance' as FAIL if deterministic domains_count >= 2.
4. For Pass 1: an Abstract claim is ONLY a discrepancy ('abstract_body_discrepancies') if it was NOT verified in ANY of the Stage 2 chunks!
5. Include ALL statistical fallacies, domain/proxy critiques, and biological/software critiques found across Stage 1 & 2.
6. In "reviewer_search_queries", provide 4 to 6 crisp 2–4 word academic search queries derived STRICTLY from the manuscript's Title and Abstract (for OpenAlex peer-reviewer search).

DETERMINISTIC GROUND-TRUTH DOSSIER:
{json.dumps(ground_truth_dossier, indent=2)}

STAGE 1 FINDINGS (Front Matter / Abstract / Intro):
{json.dumps(stage1_notes, indent=2)[:7000]}

STAGE 2 FINDINGS (Full Methods / Results / Discussion across {len(body_chunks)} chunks):
{json.dumps(stage2_notes, indent=2)[:9000]}

Output strictly valid JSON complying with the schema."""

    audit_json = call_groq_json(
        client,
        [
            {"role": "system", "content": "You are a rigorous Lead Academic Editor. Synthesize strictly valid JSON matching the schema without contradicting deterministic ground truth."},
            {"role": "user", "content": reduce_prompt},
        ],
        schema=AUDIT_STRICT_SCHEMA,
        max_tokens=4200,
    )

    reconciled = reconcile_audit_with_ground_truth(audit_json, preaudited, detected_thesis, detected_placeholders, template_obs, asset_meta)
    return reconciled, preaudited, template_obs


def reconcile_audit_with_ground_truth(
    audit: Dict[str, Any],
    preaudited: List[Dict[str, Any]],
    detected_thesis: List[str],
    detected_placeholders: List[str],
    template_obs: Dict[str, Any],
    asset_meta: Dict[str, Any],
) -> Dict[str, Any]:
    det_auth = template_obs.get("author_metadata", {})
    audit["extracted_authors"] = list(dict.fromkeys(
        (audit.get("extracted_authors") or []) + det_auth.get("authors", [])
    ))
    audit["extracted_affiliations"] = list(dict.fromkeys(
        (audit.get("extracted_affiliations") or []) + det_auth.get("affiliations", [])
    ))

    p2 = audit.setdefault("pass_2_template_and_layout", {})
    llm_sec_map = {s.get("section_name"): s for s in p2.get("structural_section_checks", [])}
    fixed_sections = []
    for pre in preaudited:
        sec_name = pre["section_name"]
        llm_entry = llm_sec_map.get(sec_name, {})
        if pre["status_hint"] == "PASS":
            fixed_sections.append({
                "section_name": sec_name,
                "detected_heading": pre["detected_heading"],
                "first_line_quote": pre["first_line_quote"],
                "status": "PASS",
            })
        else:
            fixed_sections.append({
                "section_name": sec_name,
                "detected_heading": llm_entry.get("detected_heading", pre["detected_heading"]),
                "first_line_quote": llm_entry.get("first_line_quote", pre["first_line_quote"]),
                "status": pre["status_hint"] if llm_entry.get("status") not in ("PASS", "WARN", "FAIL") else llm_entry["status"],
            })
    p2["structural_section_checks"] = fixed_sections

    if template_obs.get("multidisciplinary_domains_found") and template_obs.get("multidisciplinary_domains_count", 0) >= 2:
        p2["multidisciplinary_domains_compliance"] = "PASS"

    p2["unwanted_thesis_subheadings"] = sorted(set(
        (p2.get("unwanted_thesis_subheadings") or []) + detected_thesis
    ))
    p2["template_placeholders_detected"] = sorted(set(
        (p2.get("template_placeholders_detected") or []) + detected_placeholders
    ))

    ref_cur = template_obs.get("reference_currency", {})
    if ref_cur.get("years_detected_count", 0) >= 5 and ref_cur.get("pct_older_than_10_yrs", 0) >= 65.0:
        ref_note = (
            f"Reference currency warning: {ref_cur['pct_older_than_10_yrs']}% of indexed citation years "
            f"(median year: {ref_cur['median_year']}, range: {ref_cur['oldest_year']}–{ref_cur['newest_year']}) "
            f"are >10 years old; update literature with recent studies."
        )
        p2_notes = p2.setdefault("reference_currency_and_table_originality", [])
        if not any("median year" in n.lower() or "older" in n.lower() for n in p2_notes):
            p2_notes.append(ref_note)

    p3 = audit.setdefault("pass_3_statistical_and_mathematical_rigor", {})
    sigs = template_obs.get("scholarly_signals", {})

    if sigs.get("improper_p_values_detected"):
        p_list = p3.setdefault("p_value_reporting_issues", [])
        for pv in sigs["improper_p_values_detected"]:
            if not any(pv in existing for existing in p_list):
                p_list.append(f"Improper zero p-value detected ('{pv}'); report threshold such as 'p < 0.001'.")

    if sigs.get("pearson_directional_fallacy_hits"):
        f_list = p3.setdefault("statistical_fallacies_and_test_validity", [])
        if not any("pearson" in x.lower() or "dependent" in x.lower() for x in f_list):
            f_list.append(
                "Symmetric-vs-directional statistical fallacy: Manuscript labels variables as 'dependent' and "
                "'independent' in the context of Pearson/correlation analysis; correlation is symmetric and does not assign dependence. "
                f"Example: \"{sigs['pearson_directional_fallacy_hits'][0]}\""
            )

    if sigs.get("pseudoreplication_context_hits"):
        ps_list = p3.setdefault("pseudoreplication_flags", [])
        if not ps_list:
            ps_list.append(
                "Potential pseudoreplication detected: pixel/voxel/subsample units appear near inferential statistical tests; "
                f"clarify independent observational unit. Context: \"{sigs['pseudoreplication_context_hits'][0]}\""
            )

    p4 = audit.setdefault("pass_4_domain_and_reproducibility", {})
    if sigs.get("pejorative_terms_detected"):
        t_list = p4.setdefault("terminology_and_proxy_validity_flags", [])
        terms_str = ", ".join(sigs["pejorative_terms_detected"])
        if not any("backward" in x.lower() or "pejorative" in x.lower() for x in t_list):
            t_list.append(f"Pejorative/imprecise terminology detected ({terms_str}); replace with neutral, precise scholarly indicators.")

    if sigs.get("old_dataset_years_detected"):
        v_list = p4.setdefault("data_vintage_and_novelty_notes", [])
        yrs = ", ".join(sigs["old_dataset_years_detected"])
        if not any(y in " ".join(v_list) for y in sigs["old_dataset_years_detected"]):
            v_list.append(f"Legacy secondary dataset detected (Census/dataset year: {yrs} evaluated in {CURRENT_YEAR}); justify incremental novelty and temporal relevance.")

    p5 = audit.setdefault("pass_5_literal_text_and_captions", {})
    if sigs.get("unitalicized_taxa_xml"):
        p5["taxonomic_formatting_flags"] = sorted(set(
            (p5.get("taxonomic_formatting_flags") or []) + sigs["unitalicized_taxa_xml"]
        ))

    if sigs.get("typo_concatenation_hits"):
        p5["typographical_and_spacing_issues"] = sorted(set(
            (p5.get("typographical_and_spacing_issues") or [])
            + [f"Potential typo/concatenation: '{t}'" for t in sigs["typo_concatenation_hits"]]
        ))

    if asset_meta.get("detected_captions"):
        v_audit = []
        for cap in asset_meta["detected_captions"]:
            v_audit.append({
                "label": cap["label"],
                "placement": cap["placement"],
                "visual_present": bool(cap["has_visual"]),
                "status": "PASS" if cap["has_visual"] else "FAIL",
            })
        p5["visual_asset_audit"] = v_audit

    return audit


# ============================================================
# COI-FILTERED OPENALEX REVIEWER DISCOVERY (STRICTLY BY TITLE & ABSTRACT)
# ============================================================

def _distinctive_tokens(inst_name: str) -> set:
    words = re.findall(r"[a-z]{3,}", (inst_name or "").lower())
    return {w for w in words if w not in GENERIC_INST_TOKENS}


def is_conflict_of_interest(
    candidate_name: str,
    candidate_institution: str,
    excluded_authors: List[str],
    excluded_institutions: List[str],
) -> Optional[str]:
    cand_name_norm = normalize(candidate_name).lower()
    cand_inst_norm = normalize(candidate_institution).lower()
    cand_inst_tokens = _distinctive_tokens(cand_inst_norm)

    for auth in excluded_authors:
        auth_norm = normalize(auth).lower()
        if not auth_norm:
            continue
        if auth_norm == cand_name_norm or (len(auth_norm) >= 6 and auth_norm in cand_name_norm):
            return f"Co-author match ({auth})"
        auth_parts = auth_norm.split()
        cand_parts = cand_name_norm.split()
        if len(auth_parts) >= 2 and len(cand_parts) >= 2:
            if auth_parts[-1] == cand_parts[-1] and auth_parts[0][0] == cand_parts[0][0]:
                return f"Probable author name match ({auth})"

    for inst in excluded_institutions:
        inst_norm = normalize(inst).lower()
        if not inst_norm:
            continue
        if inst_norm in cand_inst_norm or cand_inst_norm in inst_norm:
            return f"Same institution ({inst})"
        inst_tokens = _distinctive_tokens(inst_norm)
        if inst_tokens and cand_inst_tokens:
            overlap = inst_tokens & cand_inst_tokens
            if overlap == inst_tokens or overlap == cand_inst_tokens or len(overlap) >= 2:
                return f"Institutional COI ({inst})"

    return None


def openalex_get(url: str, params: Dict[str, Any], mailto: str = "") -> Dict[str, Any]:
    req_params = dict(params)
    if mailto:
        req_params["mailto"] = mailto
    api_key = get_secret("OPENALEX_API_KEY")
    if api_key:
        req_params["api_key"] = api_key
    res = requests.get(url, params=req_params, timeout=25)
    res.raise_for_status()
    return res.json()


def _sanitize_openalex_query(q: str) -> str:
    cleaned = re.sub(r"[,:;()\[\]{}/\\]+", " ", q or "")
    return normalize(cleaned)


def _reconstruct_openalex_abstract(inv_idx: Optional[Dict[str, List[int]]]) -> str:
    if not inv_idx or not isinstance(inv_idx, dict):
        return ""
    pos_words: List[Tuple[int, str]] = []
    for word, positions in inv_idx.items():
        if isinstance(positions, list):
            for p in positions:
                pos_words.append((p, word))
    pos_words.sort(key=lambda x: x[0])
    return " ".join(w for _, w in pos_words)


def _extract_title_abstract_profile(
    title: str,
    abstract_text: str,
    extracted_keywords: List[str],
    llm_search_queries: List[str],
) -> Dict[str, Any]:
    title_words = [
        w.lower() for w in re.findall(r"\b[A-Za-z][A-Za-z\-]{2,}\b", title or "")
        if w.lower() not in ACADEMIC_STOPWORDS
    ]
    abs_words = [
        w.lower() for w in re.findall(r"\b[A-Za-z][A-Za-z\-]{2,}\b", abstract_text or "")
        if w.lower() not in ACADEMIC_STOPWORDS
    ]

    # Frequency count across Title (weighted x3) and Abstract
    freq: Dict[str, int] = {}
    for w in title_words:
        freq[w] = freq.get(w, 0) + 3
    for w in abs_words:
        freq[w] = freq.get(w, 0) + 1

    # Significant bigrams from Title and Abstract
    bigrams: Dict[str, int] = {}
    for seq, weight in [(title_words, 3), (abs_words, 1)]:
        for i in range(len(seq) - 1):
            bg = f"{seq[i]} {seq[i+1]}"
            bigrams[bg] = bigrams.get(bg, 0) + weight

    top_terms = [k for k, _ in sorted(freq.items(), key=lambda x: x[1], reverse=True)[:18]]
    top_bigrams = [k for k, v in sorted(bigrams.items(), key=lambda x: x[1], reverse=True) if v >= 2 or k in " ".join(title_words)][:8]

    # Build crisp OpenAlex queries derived strictly from Title + Abstract + Keywords
    queries: List[Tuple[str, str]] = []
    seen_q: Set[str] = set()

    def add_q(q_raw: str, label: str):
        q_clean = _sanitize_openalex_query(q_raw)
        if q_clean and len(q_clean.split()) >= 2 and q_clean.lower() not in seen_q:
            seen_q.add(q_clean.lower())
            queries.append((q_clean, label))

    for lq in (llm_search_queries or [])[:6]:
        add_q(lq, lq)

    if len(title_words) >= 2:
        add_q(" ".join(title_words[:4]), "Title Core Focus")
    if len(top_terms) >= 3:
        add_q(" ".join(top_terms[:3]), "Title & Abstract Core Terms")
    for bg in top_bigrams[:3]:
        add_q(bg, bg)
    for kw in (extracted_keywords or [])[:4]:
        add_q(kw, kw)
    if not queries and title_words:
        queries.append((title_words[0], title_words[0]))

    return {
        "title_terms": set(title_words),
        "core_terms": set(top_terms),
        "core_bigrams": set(top_bigrams),
        "queries": queries[:7],
    }


def _score_work_against_profile(work: Dict[str, Any], profile: Dict[str, Any]) -> Tuple[int, List[str]]:
    w_title = (work.get("display_name") or "").lower()
    w_abs = _reconstruct_openalex_abstract(work.get("abstract_inverted_index")).lower()
    w_topics = " ".join((t.get("display_name") or "") for t in (work.get("topics") or [])).lower()
    combined = f"{w_title} {w_abs} {w_topics}"
    work_tokens = set(re.findall(r"\b[a-z][a-z\-]{2,}\b", combined))

    matched_terms: List[str] = []
    score = 0

    for t in profile["title_terms"]:
        if t in work_tokens:
            score += 3 if t in w_title else 2
            matched_terms.append(t)

    for t in profile["core_terms"]:
        if t in work_tokens and t not in profile["title_terms"]:
            score += 2 if t in w_title else 1
            matched_terms.append(t)

    for bg in profile["core_bigrams"]:
        if bg in combined:
            score += 4
            if bg not in matched_terms:
                matched_terms.append(bg)

    return score, list(dict.fromkeys(matched_terms))[:6]


def _classify_authorship_region(
    inst_name: str,
    country_code: str,
    raw_affiliations: List[str],
    geo_city: str = "",
    geo_region: str = "",
) -> Set[str]:
    text = " ".join([inst_name, geo_city, geo_region, " ".join(raw_affiliations or [])]).lower()
    regions: Set[str] = set()

    is_assam = any(term in text for term in ASSAM_TERMS)
    is_ne_other = any(term in text for term in NORTHEAST_NON_ASSAM_TERMS)
    is_india = (country_code or "").lower() == "in" or "india" in text or is_assam or is_ne_other

    if is_assam:
        regions.add("Assam")
        regions.add("Northeast India")
    elif is_ne_other:
        regions.add("Northeast India")
    if is_india:
        regions.add("India")
    return regions


def _fetch_openalex_works_for_region(
    query: str,
    region: str,
    mailto: str = "",
) -> List[Dict[str, Any]]:
    year_cutoff = CURRENT_YEAR - 9
    aff_filter = ""
    per_page = 35

    if region == "Assam":
        aff_filter = ",raw_affiliation_strings.search:(Assam OR Guwahati OR Tezpur OR Silchar OR Dibrugarh OR Jorhat)"
        per_page = 25
    elif region == "Northeast India":
        aff_filter = (
            ",raw_affiliation_strings.search:(Meghalaya OR Shillong OR Manipur OR Imphal OR "
            "Mizoram OR Aizawl OR Nagaland OR Kohima OR Tripura OR Agartala OR Sikkim OR "
            "Gangtok OR Arunachal OR Itanagar OR NEHU)"
        )
        per_page = 25
    else:
        aff_filter = ",authorships.institutions.country_code:IN"
        per_page = 40

    # Tier 1: Strict Title + Abstract search on OpenAlex
    try:
        data = openalex_get(
            OPENALEX_URL,
            {
                "filter": f"title_and_abstract.search:{query},publication_year:>{year_cutoff}{aff_filter}",
                "per-page": per_page,
            },
            mailto=mailto,
        )
        results = data.get("results", [])
        if results:
            return results
    except Exception:
        pass

    # Tier 2: Broad OpenAlex relevance search with the same geographic/year filter
    try:
        data = openalex_get(
            OPENALEX_URL,
            {
                "search": query,
                "filter": f"publication_year:>{year_cutoff}{aff_filter.lstrip(',')}",
                "per-page": per_page,
            },
            mailto=mailto,
        )
        return data.get("results", [])
    except Exception:
        return []


def reviewer_discovery_report(
    title: str,
    abstract_text: str,
    extracted_keywords: List[str],
    llm_search_queries: List[str],
    excluded_authors: List[str],
    excluded_institutions: List[str],
    mailto: str = "",
    max_per_region: int = 4,
) -> Tuple[Dict[str, List[Dict[str, Any]]], Dict[str, Any]]:
    profile = _extract_title_abstract_profile(title, abstract_text, extracted_keywords, llm_search_queries)
    search_queries = profile["queries"]

    # Collect candidates per target region, scoring strictly against Title + Abstract
    raw_pool: Dict[str, Dict[str, Dict[str, Any]]] = {
        "Assam": {},
        "Northeast India": {},
        "India": {},
    }
    total_coi_excluded = 0

    # Query Assam & NE first, plus India-wide (India-wide also feeds NE/Assam matches if present)
    query_targets = [
        ("Assam", "Assam"),
        ("Northeast India", "Northeast India"),
        ("India", "India"),
    ]

    for query_target_region, _ in query_targets:
        for q_str, q_label in search_queries[:5]:
            works = _fetch_openalex_works_for_region(q_str, query_target_region, mailto=mailto)
            for work in works:
                rel_score, matched_terms = _score_work_against_profile(work, profile)
                # Require genuine Title/Abstract alignment
                if rel_score < 3 or len(matched_terms) < 2:
                    continue

                year = work.get("publication_year") or 0
                w_title = work.get("display_name") or ""
                doi = work.get("doi") or ""

                for authorship in work.get("authorships", []):
                    author = authorship.get("author") or {}
                    author_id = author.get("id")
                    author_name = author.get("display_name")
                    if not author_id or not author_name:
                        continue

                    raw_affs = authorship.get("raw_affiliation_strings") or []
                    institutions = authorship.get("institutions") or []
                    if not institutions and raw_affs:
                        institutions = [{"display_name": raw_affs[0], "country_code": "IN", "geo": {}}]

                    for inst in institutions:
                        inst_name = inst.get("display_name") or (raw_affs[0] if raw_affs else "Unspecified Institution")
                        geo = inst.get("geo") or {}
                        matched_regions = _classify_authorship_region(
                            inst_name=inst_name,
                            country_code=inst.get("country_code") or "",
                            raw_affiliations=raw_affs,
                            geo_city=geo.get("city") or "",
                            geo_region=geo.get("region") or "",
                        )
                        if not matched_regions:
                            continue

                        if is_conflict_of_interest(author_name, inst_name, excluded_authors, excluded_institutions):
                            total_coi_excluded += 1
                            continue

                        pub_record = {
                            "year": year,
                            "title": w_title,
                            "doi": doi,
                            "relevance_score": rel_score,
                            "matched_terms": matched_terms,
                        }

                        for reg in matched_regions:
                            bucket = raw_pool[reg]
                            if author_id not in bucket:
                                bucket[author_id] = {
                                    "author_id": author_id,
                                    "name": author_name,
                                    "institution": inst_name,
                                    "raw_affiliation": raw_affs[0] if raw_affs else inst_name,
                                    "match_type": f"Title/Abstract match: {', '.join(matched_terms[:4])}",
                                    "best_relevance": rel_score,
                                    "matched_terms": set(matched_terms),
                                    "recent_pubs": [],
                                    "_seen_titles": set(),
                                }
                            cand = bucket[author_id]
                            cand["best_relevance"] = max(cand["best_relevance"], rel_score)
                            cand["matched_terms"].update(matched_terms)
                            cand["match_type"] = f"Title/Abstract match: {', '.join(sorted(cand['matched_terms'])[:4])}"
                            if w_title not in cand["_seen_titles"]:
                                cand["_seen_titles"].add(w_title)
                                cand["recent_pubs"].append(pub_record)

            # Stop early for Assam / NE once enough strong candidates exist
            if query_target_region != "India" and len(raw_pool[query_target_region]) >= max_per_region * 2:
                break

    # Final selection: Assam -> Northeast India -> India (deduplicated across lists so each list is distinct)
    output: Dict[str, List[Dict[str, Any]]] = {"India": [], "Northeast India": [], "Assam": []}
    assigned_authors: Set[str] = set()

    def rank_candidates(cands_dict: Dict[str, Dict[str, Any]]) -> List[Dict[str, Any]]:
        vals = list(cands_dict.values())
        for c in vals:
            c["recent_pubs"].sort(key=lambda p: (p["relevance_score"], p["year"]), reverse=True)
            c.pop("_seen_titles", None)
            c["matched_terms"] = sorted(c["matched_terms"])
        vals.sort(
            key=lambda c: (
                c["best_relevance"],
                len(c["recent_pubs"]),
                c["recent_pubs"][0]["year"] if c["recent_pubs"] else 0,
            ),
            reverse=True,
        )
        return vals

    for reg in ["Assam", "Northeast India", "India"]:
        ranked = rank_candidates(raw_pool[reg])
        selected: List[Dict[str, Any]] = []
        for cand in ranked:
            if cand["author_id"] in assigned_authors:
                continue
            selected.append(cand)
            assigned_authors.add(cand["author_id"])
            if len(selected) >= max_per_region:
                break
        output[reg] = selected

    coi_summary = {
        "excluded_authors": excluded_authors,
        "excluded_institutions": excluded_institutions,
        "coi_candidates_filtered": total_coi_excluded,
        "search_queries_used": [q for q, _ in search_queries],
    }
    return output, coi_summary


# ============================================================
# BLIND COPY BUILDER (REMOVES AUTHOR DETAILS ONLY; REST UNCHANGED)
# ============================================================

def _replace_regex_in_paragraph(para: Paragraph, pattern: re.Pattern, replacement: str) -> None:
    full_text = para.text or ""
    if not pattern.search(full_text):
        return
    # Fast path: single-run matches preserve run-level formatting
    replaced_in_run = False
    for r in para.runs:
        if pattern.search(r.text or ""):
            r.text = pattern.sub(replacement, r.text)
            replaced_in_run = True
    # Slow path: match spanned multiple runs
    if pattern.search(para.text or "") or not replaced_in_run:
        new_full = pattern.sub(replacement, para.text or "")
        if para.runs:
            para.runs[0].text = new_full
            for r in para.runs[1:]:
                r.text = ""
        else:
            para.text = new_full


def _clear_paragraph(para: Paragraph) -> None:
    for r in para.runs:
        r.text = ""
    for elem in list(para._element):
        tag = elem.tag.split("}")[-1]
        if tag in ("r", "hyperlink"):
            para._element.remove(elem)


def _remove_paragraph_element(para: Paragraph) -> None:
    parent = para._element.getparent()
    if parent is not None:
        parent.remove(para._element)


def _is_title_paragraph(txt: str, manuscript_title: str) -> bool:
    norm_txt = normalize(txt).lower()
    norm_title = normalize(manuscript_title).lower()
    if not norm_txt or not norm_title:
        return False
    if norm_txt in norm_title or norm_title in norm_txt:
        return True
    t_words = {w for w in re.findall(r"[a-z0-9]{3,}", norm_title)}
    p_words = {w for w in re.findall(r"[a-z0-9]{3,}", norm_txt)}
    if not t_words or not p_words:
        return False
    overlap = len(t_words & p_words) / max(len(t_words), 1)
    return overlap >= 0.55


def _has_author_or_affiliation_signal(
    txt: str,
    extracted_authors: List[str],
    extracted_affiliations: List[str],
) -> bool:
    clean = normalize(txt)
    if not clean:
        return False
    lower = clean.lower()

    if re.search(r"\b[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}\b", clean, re.I):
        return True
    if re.search(r"(?:https?://)?orcid\.org/\d{4}-\d{4}-\d{4}-[\dX]{4}", clean, re.I):
        return True
    if re.match(
        r"^(?:\*|†|‡)?\s*(?:corresponding\s+author|correspondence(?:\s+to)?|author\s+contributions?|authors?['’]?\s+contributions?|e-?mail|affiliations?|orcid)\b",
        lower,
    ):
        return True

    for auth in extracted_authors:
        a = normalize(auth).lower()
        if len(a) >= 4 and re.search(rf"(?<!\w){re.escape(a)}(?!\w)", lower):
            return True

    for inst in extracted_affiliations:
        i_low = normalize(inst).lower()
        if len(i_low) >= 6 and i_low in lower:
            return True
        inst_toks = _distinctive_tokens(i_low)
        txt_toks = _distinctive_tokens(lower)
        if len(inst_toks) >= 2 and len(inst_toks & txt_toks) >= 2:
            return True

    if re.search(
        r"\b(?:university|institute|college|department\s+of|faculty\s+of|school\s+of|centre\s+for|center\s+for|academy|laboratory|iit\b|nit\b|iiit\b|aiims\b|csir\b|icar\b)\b",
        lower,
    ):
        return True

    return False


def blind_copy_docx(
    original_bytes: bytes,
    manuscript_title: str = "",
    extracted_authors: Optional[List[str]] = None,
    extracted_affiliations: Optional[List[str]] = None,
) -> bytes:
    authors_list = [a for a in (extracted_authors or []) if normalize(a)]
    affils_list = [a for a in (extracted_affiliations or []) if normalize(a)]

    doc = Document(io.BytesIO(original_bytes))
    props = doc.core_properties
    props.author = ""
    props.last_modified_by = ""
    props.comments = ""

    email_re = re.compile(r"\b[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}\b", re.I)
    orcid_re = re.compile(r"(?:https?://)?orcid\.org/\d{4}-\d{4}-\d{4}-[\dX]{4}", re.I)
    explicit_author_block_re = re.compile(
        r"^(?:\*|†|‡)?\s*(?:corresponding\s+author|correspondence(?:\s+to)?|author\s+contributions?|authors?['’]?\s+contributions?|orcid\s*[:\-])\b",
        re.I,
    )
    journal_meta_re = re.compile(
        r"^(?:received|accepted|published|issn|doi|keywords?|running\s+title|original\s+(?:research\s+)?article|research\s+article|research\s+paper)\b",
        re.I,
    )

    in_front_matter = True
    seen_title = False

    body_children = list(doc.element.body.iterchildren())
    for child in body_children:
        if isinstance(child, CT_P):
            p = Paragraph(child, doc)
            txt = normalize(p.text)

            # Detect start of body (Abstract / Keywords / Introduction / implicit abstract)
            if in_front_matter and txt:
                if (
                    re.match(r"^(?:(?:\d+\.?\s+)?(?:abstract|executive\s+summary|introduction|keywords?))\b", txt, re.I)
                    or (
                        word_count(txt) >= 35
                        and not _has_author_or_affiliation_signal(txt, authors_list, affils_list)
                        and not _is_title_paragraph(txt, manuscript_title)
                    )
                ):
                    in_front_matter = False

            if in_front_matter:
                if not txt:
                    continue
                if _is_title_paragraph(txt, manuscript_title):
                    seen_title = True
                    continue
                if journal_meta_re.match(txt):
                    continue

                is_byline_after_title = (
                    seen_title
                    and word_count(txt) <= 25
                    and (
                        bool(re.search(r"[*†‡§\d]", txt))
                        or "," in txt
                        or " and " in txt.lower()
                    )
                )
                if _has_author_or_affiliation_signal(txt, authors_list, affils_list) or is_byline_after_title:
                    _remove_paragraph_element(p)
                    continue
            else:
                if explicit_author_block_re.match(txt):
                    _remove_paragraph_element(p)
                    continue
                _replace_regex_in_paragraph(p, email_re, "[REDACTED]")
                _replace_regex_in_paragraph(p, orcid_re, "[REDACTED]")

            # Clean author-identifying textboxes inside front-matter paragraphs
            for txbx in child.findall(".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}txbxContent"):
                for tp_elem in list(txbx.findall(".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p")):
                    tp = Paragraph(tp_elem, doc)
                    t_txt = normalize(tp.text)
                    if not t_txt:
                        continue
                    if (in_front_matter and not _is_title_paragraph(t_txt, manuscript_title) and _has_author_or_affiliation_signal(t_txt, authors_list, affils_list)) or explicit_author_block_re.match(t_txt):
                        _clear_paragraph(tp)
                    else:
                        _replace_regex_in_paragraph(tp, email_re, "[REDACTED]")
                        _replace_regex_in_paragraph(tp, orcid_re, "[REDACTED]")

        elif isinstance(child, CT_Tbl):
            tbl = Table(child, doc)
            for row in list(tbl.rows):
                for cell in row.cells:
                    for cp in list(cell.paragraphs):
                        ctxt = normalize(cp.text)
                        if not ctxt:
                            continue
                        if in_front_matter and not _is_title_paragraph(ctxt, manuscript_title) and _has_author_or_affiliation_signal(ctxt, authors_list, affils_list):
                            _clear_paragraph(cp)
                        elif explicit_author_block_re.match(ctxt):
                            _clear_paragraph(cp)
                        else:
                            _replace_regex_in_paragraph(cp, email_re, "[REDACTED]")
                            _replace_regex_in_paragraph(cp, orcid_re, "[REDACTED]")

    # Clean Headers & Footers across all sections without altering non-author headers/page numbers
    for section in doc.sections:
        for hf in [
            section.header,
            section.footer,
            section.first_page_header,
            section.first_page_footer,
            section.even_page_header,
            section.even_page_footer,
        ]:
            if not hf:
                continue
            for p in list(hf.paragraphs):
                htxt = normalize(p.text)
                if not htxt:
                    continue
                if _has_author_or_affiliation_signal(htxt, authors_list, affils_list):
                    _clear_paragraph(p)
                else:
                    _replace_regex_in_paragraph(p, email_re, "[REDACTED]")
                    _replace_regex_in_paragraph(p, orcid_re, "[REDACTED]")
            for tbl in hf.tables:
                for row in tbl.rows:
                    for cell in row.cells:
                        for p in list(cell.paragraphs):
                            htxt = normalize(p.text)
                            if htxt and _has_author_or_affiliation_signal(htxt, authors_list, affils_list):
                                _clear_paragraph(p)
                            else:
                                _replace_regex_in_paragraph(p, email_re, "[REDACTED]")
                                _replace_regex_in_paragraph(p, orcid_re, "[REDACTED]")

    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()


# ============================================================
# MARKDOWN PREVIEW & COMPREHENSIVE ALL-COMMENTS DOCX BUILDERS
# ============================================================

def generate_markdown_audit_report(
    audit: Dict[str, Any],
    reviewers: Dict[str, List[Dict[str, Any]]],
    coi_summary: Dict[str, Any],
) -> str:
    verd = audit.get("editorial_verdict", {})
    p1 = audit.get("pass_1_cross_section_integrity", {})
    p2 = audit.get("pass_2_template_and_layout", {})
    p3 = audit.get("pass_3_statistical_and_mathematical_rigor", {})
    p4 = audit.get("pass_4_domain_and_reproducibility", {})
    p5 = audit.get("pass_5_literal_text_and_captions", {})

    reqs = verd.get("actionable_revision_requirements", [])
    reqs_md = "\n".join([f"- {r}" for r in reqs]) if reqs else "- No mandatory revisions requested."

    unwanted = p2.get("unwanted_thesis_subheadings", [])
    unwanted_str = ", ".join(unwanted) if unwanted else "None detected"
    placeholders = p2.get("template_placeholders_detected", [])
    placeholders_str = ", ".join(placeholders) if placeholders else "None"
    ref_orig = p2.get("reference_currency_and_table_originality", [])
    ref_orig_md = "\n".join([f"- {x}" for x in ref_orig]) if ref_orig else "- Literature currency and table originality satisfactory."

    disc = p1.get("abstract_body_discrepancies", [])
    disc_str = "\n".join([f"- {d}" for d in disc]) if disc else "- All Abstract claims are empirically corroborated in Methods and Results."

    p_issues = p3.get("p_value_reporting_issues", [])
    p_issues_str = "; ".join(p_issues) if p_issues else "None (properly formatted)"
    pseudo = p3.get("pseudoreplication_flags", [])
    pseudo_str = "; ".join(pseudo) if pseudo else "None identified"
    stat_fallacies = p3.get("statistical_fallacies_and_test_validity", [])
    stat_fallacies_md = "\n".join([f"- {s}" for s in stat_fallacies]) if stat_fallacies else "- No symmetric/directional or test-validity fallacies identified."

    vintage_notes = p4.get("data_vintage_and_novelty_notes", [])
    term_proxy = p4.get("terminology_and_proxy_validity_flags", [])
    overinterp = p4.get("conceptual_overinterpretation_flags", [])
    soft_map = p4.get("software_novelty_and_mapping_notes", [])
    edge_cases = p4.get("algorithmic_edge_cases_notes", [])

    enum_issues = p5.get("enumeration_figure_inconsistencies", [])
    enum_str = "; ".join(enum_issues) if enum_issues else "Consistent with figures"
    taxa = p5.get("taxonomic_formatting_flags", [])
    typos = p5.get("typographical_and_spacing_issues", [])
    taxa_str = "; ".join(taxa + typos) if (taxa or typos) else "Standard binomial nomenclature and formatting"

    rev_lines = []
    excluded_insts = coi_summary.get("excluded_institutions", [])
    if excluded_insts:
        rev_lines.append(f"*COI Protection Active — Excluded author affiliations:* {', '.join(excluded_insts)}\n")
    for reg, cands in reviewers.items():
        rev_lines.append(f"#### {reg}")
        if not cands:
            rev_lines.append("- No COI-free Title/Abstract-matched OpenAlex profiles identified for this region.")
        for c in cands:
            pubs = c.get("recent_pubs", [])
            top_pub = f" — Latest match: *{pubs[0]['title']}* ({pubs[0]['year']})" if pubs else ""
            rev_lines.append(f"- **{c.get('name')}** ({c.get('institution')}): {c.get('match_type')} [{len(pubs)} matching works]{top_pub}")

    return f"""# MRJ Academic Pre-Screening & Editorial Audit Report

**Manuscript Title:** {audit.get('manuscript_title', 'Not specified')}  
**Detected Authors:** {', '.join(audit.get('extracted_authors', [])) or 'Not extracted'}  
**Detected Affiliations:** {', '.join(audit.get('extracted_affiliations', [])) or 'Not extracted'}  
**Decision:** {verd.get('decision', 'Major Revisions')}

### Executive Summary & Actionable Revision Requirements
{verd.get('executive_summary', '')}

{reqs_md}

### 1. Cross-Section Consistency & Phantom Claim Audit
- **Abstract vs. Body Discrepancies:**
{disc_str}

### 2. Template Compliance, Structural Audit & Reference Currency
- **Journal Style Compliance:** {p2.get('journal_style_compliance', 'FAIL')}
- **Unwanted Thesis Subheadings:** {unwanted_str}
- **Template Placeholders Detected:** {placeholders_str}
- **Multidisciplinary Domains Statement:** {p2.get('multidisciplinary_domains_compliance', 'FAIL')}
- **Abstract Quality & Word Count:** {p2.get('abstract_quality_and_word_count', 'N/A')}
- **Reference Currency & Table Originality:**
{ref_orig_md}

### 3. Statistical, Mathematical & Sampling Unit Rigor
- **P-Value Reporting Issues:** {p_issues_str}
- **Pseudoreplication / Observational Unit Flags:** {pseudo_str}
- **Statistical Fallacies & Test Validity (Pearson/Chi-Square/ANOVA):**
{stat_fallacies_md}
- **Statistical Test Completeness:** {p3.get('statistical_test_completeness', 'N/A')}

### 4. Domain Validity, Conceptual Framing & Reproducibility
- **Data Vintage & Novelty:** {'; '.join(vintage_notes) if vintage_notes else 'Current/appropriate'}
- **Terminology & Proxy Variable Validity:** {'; '.join(term_proxy) if term_proxy else 'No imprecise terminology or invalid proxies flagged'}
- **Conceptual / Biological Over-Interpretation:** {'; '.join(overinterp) if overinterp else 'Claims proportional to empirical metrics'}
- **Software Novelty & Mathematical Mapping:** {'; '.join(soft_map) if soft_map else 'N/A or adequately specified'}
- **Algorithmic Edge Cases & Versions:** {'; '.join(edge_cases) if edge_cases else 'None'} | Versions: {p4.get('software_dependency_versions', 'N/A')}

### 5. Text-Level Accuracy, Taxonomic Formatting & Visual Captions
- **Enumeration / Text-Figure Inconsistencies:** {enum_str}
- **Taxonomic Formatting & Typos:** {taxa_str}

### 6. COI-Filtered Verified Reviewer Discovery (OpenAlex — Title & Abstract Matched)
{chr(10).join(rev_lines)}
"""


def generate_docx_report(
    audit: Dict[str, Any],
    reviewers: Dict[str, List[Dict[str, Any]]],
    coi_summary: Dict[str, Any],
) -> bytes:
    doc = Document()
    doc.add_heading("MRJ All-Comments Editorial Audit & Reviewer Report", 0)
    doc.add_paragraph(f"Manuscript Title: {audit.get('manuscript_title', 'Not specified')}")
    if audit.get("extracted_authors"):
        doc.add_paragraph(f"Detected Authors: {', '.join(audit['extracted_authors'])}")
    if audit.get("extracted_affiliations"):
        doc.add_paragraph(f"Author Affiliations (Excluded from Reviewer Match): {', '.join(audit['extracted_affiliations'])}")

    verd = audit.get("editorial_verdict", {})
    doc.add_paragraph(f"Editorial Decision: {verd.get('decision', 'Under Review')}")
    doc.add_paragraph(f"Executive Summary: {verd.get('executive_summary', '')}")

    reqs = verd.get("actionable_revision_requirements", [])
    if reqs:
        doc.add_heading("Actionable Revision Requirements for Authors", level=2)
        for r in reqs:
            doc.add_paragraph(f"• {r}")

    # Pass 1
    doc.add_heading("1. Cross-Section Content Integrity & Discrepancy Audit", level=1)
    p1 = audit.get("pass_1_cross_section_integrity", {})
    doc.add_paragraph("Claims Extracted from Abstract:")
    for item in p1.get("abstract_claimed_items", []) or ["None extracted"]:
        doc.add_paragraph(f"• {item}")
    doc.add_paragraph("Claims Verified in Methods / Results:")
    for item in p1.get("verified_in_body_items", []) or ["None verified"]:
        doc.add_paragraph(f"• {item}")
    doc.add_paragraph("Abstract vs. Body Discrepancies:")
    disc = p1.get("abstract_body_discrepancies", [])
    if disc:
        for d in disc:
            doc.add_paragraph(f"• {d}")
    else:
        doc.add_paragraph("• All Abstract claims are corroborated by empirical data in Methods and Results.")

    # Pass 2
    doc.add_heading("2. Template Compliance, Structural Audit & Reference Currency", level=1)
    p2 = audit.get("pass_2_template_and_layout", {})
    doc.add_paragraph(f"Journal Style Compliance: {p2.get('journal_style_compliance', 'N/A')}")
    doc.add_paragraph(f"Multidisciplinary Domains Statement: {p2.get('multidisciplinary_domains_compliance', 'N/A')}")
    doc.add_paragraph(f"Abstract Word Count & Quality: {p2.get('abstract_quality_and_word_count', 'N/A')}")
    doc.add_paragraph(f"Unwanted Thesis Subheadings: {', '.join(p2.get('unwanted_thesis_subheadings') or []) or 'None detected'}")
    doc.add_paragraph(f"Template Placeholders Detected: {', '.join(p2.get('template_placeholders_detected') or []) or 'None detected'}")
    for note in p2.get("reference_currency_and_table_originality", []):
        doc.add_paragraph(f"• {note}")

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
    doc.add_heading("3. Statistical, Mathematical & Sampling Unit Rigor", level=1)
    p3 = audit.get("pass_3_statistical_and_mathematical_rigor", {})
    doc.add_paragraph(f"P-Value Issues: {', '.join(p3.get('p_value_reporting_issues') or []) or 'None detected'}")
    doc.add_paragraph(f"Pseudoreplication Flags: {', '.join(p3.get('pseudoreplication_flags') or []) or 'None flagged'}")
    for sf in p3.get("statistical_fallacies_and_test_validity", []):
        doc.add_paragraph(f"• {sf}")
    doc.add_paragraph(f"Statistical Completeness: {p3.get('statistical_test_completeness', 'N/A')}")

    # Pass 4
    doc.add_heading("4. Domain Validity, Conceptual Framing & Reproducibility", level=1)
    p4 = audit.get("pass_4_domain_and_reproducibility", {})
    doc.add_paragraph(f"Computational / Software Paper: {'Yes' if p4.get('is_computational_or_algorithm_paper') else 'No'}")
    for label, key in [
        ("Data Vintage & Novelty", "data_vintage_and_novelty_notes"),
        ("Terminology / Proxy Validity", "terminology_and_proxy_validity_flags"),
        ("Conceptual Over-Interpretation", "conceptual_overinterpretation_flags"),
        ("Software Novelty & Mapping", "software_novelty_and_mapping_notes"),
        ("Algorithmic Edge Cases", "algorithmic_edge_cases_notes"),
    ]:
        items = p4.get(key) or []
        if items:
            for item in items:
                doc.add_paragraph(f"• [{label}] {item}")
    doc.add_paragraph(f"Software Versions: {p4.get('software_dependency_versions', 'None specified')}")
    doc.add_paragraph(f"Benchmarks / Scalability: {p4.get('benchmarks_and_scalability', 'None specified')}")

    # Pass 5
    doc.add_heading("5. Text-Level Accuracy & Visual Captions", level=1)
    p5 = audit.get("pass_5_literal_text_and_captions", {})
    all_p5_flags = (
        p5.get("enumeration_figure_inconsistencies", [])
        + p5.get("taxonomic_formatting_flags", [])
        + p5.get("typographical_and_spacing_issues", [])
    )
    if all_p5_flags:
        for fl in all_p5_flags:
            doc.add_paragraph(f"• {fl}")
    else:
        doc.add_paragraph("• No enumeration, taxonomic formatting, or typographical flags detected.")

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

    # Reviewers (India, Northeast India, Assam matched strictly to Title & Abstract)
    doc.add_heading("6. COI-Filtered Verified Reviewer Discovery (India, Northeast India, Assam)", level=1)
    if coi_summary.get("search_queries_used"):
        doc.add_paragraph(f"Title & Abstract Search Queries Used: {', '.join(coi_summary['search_queries_used'])}")
    if coi_summary.get("excluded_institutions"):
        doc.add_paragraph(f"Excluded for Conflict of Interest: {', '.join(coi_summary['excluded_institutions'])}")

    for reg in ["India", "Northeast India", "Assam"]:
        cands = reviewers.get(reg, [])
        doc.add_heading(reg, level=2)
        if not cands:
            doc.add_paragraph("No COI-free matching reviewers identified for the Title & Abstract.")
            continue
        r_tbl = doc.add_table(rows=1, cols=4)
        r_tbl.style = "Table Grid"
        rh = r_tbl.rows[0].cells
        rh[0].text, rh[1].text, rh[2].text, rh[3].text = (
            "Reviewer Name",
            "Institution / Affiliation",
            "Title & Abstract Alignment",
            "Matching Publications",
        )
        for c in cands:
            row = r_tbl.add_row().cells
            row[0].text = c.get("name", "")
            row[1].text = c.get("raw_affiliation") or c.get("institution", "")
            row[2].text = c.get("match_type", "")
            pubs = c.get("recent_pubs", [])[:2]
            row[3].text = "\n".join(
                f"({p.get('year', '')}) {p.get('title', '')} {p.get('doi', '')}".strip()
                for p in pubs
            ) or "N/A"

    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()


# ============================================================
# STREAMLIT UI
# ============================================================

st.set_page_config(page_title="MRJ Manuscript Auditor (Hardened 5-Pass Engine)", page_icon="📄", layout="wide")
st.title("📄 MRJ Manuscript Quality Auditor & Pre-Screening Engine")
st.caption("Full-manuscript zero-truncation 5-Pass editorial audit: Template compliance, cross-section verification, statistical/sampling rigor, domain & proxy validity, algorithmic reproducibility, and COI-safe OpenAlex reviewer discovery matched to Title & Abstract.")

with st.sidebar:
    st.header("⚙️ Editorial Settings")
    groq_api_key = st.text_input("Groq API Key", value=get_secret("GROQ_API_KEY"), type="password")
    openalex_mailto = st.text_input("OpenAlex Mailto Email", value=get_secret("OPENALEX_MAILTO", "editor@mrjournal.org"))
    tier_choice = st.radio(
        "Groq API Throughput Profile",
        ["Safe Chunked (Free Tier / 8K TPM)", "Fast Full-Context (Developer Tier / 250K TPM)"],
        index=0,
        help="Safe Chunked processes 100% of the manuscript in ~9.5k-char batches with TPM backoff so nothing is truncated.",
    )
    free_tier_mode = tier_choice.startswith("Safe Chunked")

    st.markdown("---")
    st.markdown("**Hardened 5-Pass Audit Core:**")
    st.markdown("1. **Pass 1:** Cross-section integrity across full Methods & Results.")
    st.markdown("2. **Pass 2:** Flexible section numbering, domain lists, placeholders & reference currency.")
    st.markdown("3. **Pass 3:** Pearson directional fallacy, pixel pseudoreplication, Chi-square/ANOVA validity.")
    st.markdown("4. **Pass 4:** Data vintage, proxy/terminology audit, over-interpretation & algorithmic edge cases.")
    st.markdown("5. **Pass 5:** XML italicization check, typos & full table/figure inventory.")

uploaded_file = st.file_uploader("Upload Manuscript (.docx or .pdf)", type=["docx", "pdf"])

if uploaded_file and st.button("🚀 Run Full-Text 5-Pass MRJ Audit", type="primary"):
    if not groq_api_key:
        st.error("Please provide a valid Groq API Key.")
        st.stop()

    try:
        status_box = st.status("Initializing Full-Document Audit...", expanded=True)
        with status_box:
            st.write("Extracting complete manuscript text, interleaved tables, and visual assets...")
            raw_text, asset_meta = extract_text_and_assets(uploaded_file)
            if not raw_text.strip():
                st.error("Could not extract readable text from the uploaded document.")
                st.stop()

            st.write(f"Parsed {word_count(raw_text):,} words | {asset_meta['total_tables']} tables | {asset_meta['total_images']} embedded images.")
            client = Groq(api_key=groq_api_key)
            audit_result, _, template_obs = run_editorial_audit(
                raw_text,
                asset_meta,
                client,
                free_tier_mode=free_tier_mode,
                progress_cb=lambda msg: st.write(msg),
            )

            st.write("Discovering COI-filtered reviewers from India, Northeast India & Assam strictly matched to Title & Abstract...")
            detected_title = audit_result.get("manuscript_title", "")
            abstract_text = template_obs.get("abstract_text", "")
            keywords = template_obs.get("keywords_detected", [])
            llm_queries = audit_result.get("reviewer_search_queries", [])
            reviewers, coi_summary = reviewer_discovery_report(
                title=detected_title,
                abstract_text=abstract_text,
                extracted_keywords=keywords,
                llm_search_queries=llm_queries,
                excluded_authors=audit_result.get("extracted_authors", []),
                excluded_institutions=audit_result.get("extracted_affiliations", []),
                mailto=openalex_mailto,
            )

            st.write("Generating All-Comments Report (.docx) and Author-Blind Reviewer Copy (.docx)...")
            md_report = generate_markdown_audit_report(audit_result, reviewers, coi_summary)
            orig_bytes = uploaded_file.getvalue()
            if uploaded_file.name.lower().endswith(".docx"):
                blind_bytes = blind_copy_docx(
                    orig_bytes,
                    manuscript_title=detected_title,
                    extracted_authors=audit_result.get("extracted_authors", []),
                    extracted_affiliations=audit_result.get("extracted_affiliations", []),
                )
            else:
                blind_bytes = None
            docx_report = generate_docx_report(audit_result, reviewers, coi_summary)
            status_box.update(label="Full-Manuscript 5-Pass Audit Complete!", state="complete", expanded=False)

        st.session_state["audit"] = audit_result
        st.session_state["reviewers"] = reviewers
        st.session_state["coi_summary"] = coi_summary
        st.session_state["template_obs"] = template_obs
        st.session_state["md_report"] = md_report
        st.session_state["blind_bytes"] = blind_bytes
        st.session_state["docx_report"] = docx_report

    except Exception as e:
        st.error(f"Audit processing error: {e}")


if "audit" in st.session_state:
    audit = st.session_state["audit"]
    reviewers = st.session_state["reviewers"]
    coi_summary = st.session_state["coi_summary"]
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
        "📄 All-Comments Summary",
        "Pass 1: Cross-Section Integrity",
        "Pass 2: Template & Citations",
        "Pass 3: Statistical Rigor",
        "Pass 4: Domain & Reproducibility",
        "Pass 5: Text & Captions",
        "👥 COI-Safe Reviewers",
    ])

    with t_md:
        st.markdown(md_report)

    with t1:
        st.subheader("Pass 1: Cross-Section Content Integrity & Phantom Claim Audit")
        p1 = audit.get("pass_1_cross_section_integrity", {})
        col_a, col_b = st.columns(2)
        with col_a:
            st.markdown("**Claims Extracted from Abstract:**")
            for item in p1.get("abstract_claimed_items", []):
                st.markdown(f"- 🔬 {item}")
        with col_b:
            st.markdown("**Verified in Methods/Results:**")
            for item in p1.get("verified_in_body_items", []):
                st.markdown(f"- ✅ {item}")

        discs = p1.get("abstract_body_discrepancies", [])
        if discs:
            st.error("**Discrepancies & Phantom Claims Flagged:**")
            for d in discs:
                st.markdown(f"- ❌ {d}")
        else:
            st.success("✅ No discrepancies detected between Abstract claims and empirical data.")

    with t2:
        st.subheader("Pass 2: Template, Layout, Placeholders & Literature Currency")
        p2 = audit.get("pass_2_template_and_layout", {})
        ref_cur = template_obs.get("reference_currency", {})
        col_p1, col_p2, col_p3, col_p4 = st.columns(4)
        col_p1.metric("Journal Style Compliance", p2.get("journal_style_compliance", "N/A"))
        col_p2.metric("Multidisciplinary Domains", p2.get("multidisciplinary_domains_compliance", "N/A"))
        col_p3.metric("Abstract Words", f"~{template_obs.get('abstract_word_count', 0)}")
        col_p4.metric("Median Ref Year", ref_cur.get("median_year") or "N/A")

        for note in p2.get("reference_currency_and_table_originality", []):
            st.warning(f"📚 {note}")

        phs = p2.get("template_placeholders_detected", [])
        if phs:
            st.warning("⚠️ **Template Placeholders Detected:** " + ", ".join([f"`{p}`" for p in phs]))

        unw = p2.get("unwanted_thesis_subheadings", [])
        if unw:
            st.warning("⚠️ **Unwanted Dissertation Subheadings:** " + ", ".join([f"`{u}`" for u in unw]))

        st.dataframe(p2.get("structural_section_checks", []), use_container_width=True, hide_index=True)

    with t3:
        st.subheader("Pass 3: Statistical, Mathematical & Sampling Unit Rigor")
        p3 = audit.get("pass_3_statistical_and_mathematical_rigor", {})
        for pi in p3.get("p_value_reporting_issues", []):
            st.warning(f"⚠️ **P-Value Reporting:** {pi}")
        for ps in p3.get("pseudoreplication_flags", []):
            st.error(f"🚨 **Pseudoreplication / Observational Unit:** {ps}")
        for sf in p3.get("statistical_fallacies_and_test_validity", []):
            st.error(f"🚨 **Statistical Fallacy / Test Validity:** {sf}")
        if not (p3.get("p_value_reporting_issues") or p3.get("pseudoreplication_flags") or p3.get("statistical_fallacies_and_test_validity")):
            st.success("✅ No statistical reporting or sampling-unit fallacies detected.")
        st.info(f"**Statistical Test Completeness:** {p3.get('statistical_test_completeness', 'N/A')}")

    with t4:
        st.subheader("Pass 4: Domain Validity, Conceptual Framing & Reproducibility")
        p4 = audit.get("pass_4_domain_and_reproducibility", {})
        for dv in p4.get("data_vintage_and_novelty_notes", []):
            st.warning(f"🗓️ **Data Vintage & Novelty:** {dv}")
        for tp in p4.get("terminology_and_proxy_validity_flags", []):
            st.warning(f"🏷️ **Terminology / Proxy Variable:** {tp}")
        for co in p4.get("conceptual_overinterpretation_flags", []):
            st.error(f"🧠 **Conceptual Over-Interpretation:** {co}")

        st.write(f"**Computational/Software Tool Paper:** {'Yes' if p4.get('is_computational_or_algorithm_paper') else 'No'}")
        for sm in p4.get("software_novelty_and_mapping_notes", []):
            st.warning(f"🧩 **Software Novelty / Mapping:** {sm}")
        for ec in p4.get("algorithmic_edge_cases_notes", []):
            st.warning(f"⚙️ **Edge Case:** {ec}")
        st.write(f"**Software Versions:** {p4.get('software_dependency_versions', 'None specified')}")
        st.write(f"**Scalability & Benchmarks:** {p4.get('benchmarks_and_scalability', 'None recorded')}")

    with t5:
        st.subheader("Pass 5: Literal Text & Visual Caption Audit")
        p5 = audit.get("pass_5_literal_text_and_captions", {})
        for en in p5.get("enumeration_figure_inconsistencies", []):
            st.warning(f"⚠️ **Enumeration / Figure Mismatch:** {en}")
        for tx in p5.get("taxonomic_formatting_flags", []):
            st.warning(f"🧬 **Taxonomic Formatting:** {tx}")
        for ty in p5.get("typographical_and_spacing_issues", []):
            st.warning(f"✏️ **Typo / Spacing:** {ty}")
        st.dataframe(p5.get("visual_asset_audit", []), use_container_width=True, hide_index=True)

    with t_rev:
        st.subheader("COI-Filtered Relevant Reviewer Discovery (OpenAlex — Title & Abstract Matched)")
        if coi_summary.get("search_queries_used"):
            st.caption(f"**Title & Abstract Queries:** {' | '.join(coi_summary['search_queries_used'])}")
        if coi_summary.get("excluded_institutions") or coi_summary.get("excluded_authors"):
            st.info(
                f"🛡️ **COI Filter Active:** Excluded **{coi_summary.get('coi_candidates_filtered', 0)}** candidate matches affiliated with "
                f"{', '.join(coi_summary.get('excluded_institutions', [])) or 'manuscript institutions'}."
            )
        for region in ["India", "Northeast India", "Assam"]:
            cands = reviewers.get(region, [])
            st.markdown(f"### Region: {region}")
            if not cands:
                st.write("No COI-free Title/Abstract-matched candidates located.")
                continue
            r_rows = [
                {
                    "Candidate Name": c.get("name"),
                    "Affiliation": c.get("raw_affiliation") or c.get("institution"),
                    "Title & Abstract Alignment": c.get("match_type"),
                    "Matching Works": len(c.get("recent_pubs", [])),
                    "Top Matching Paper": (
                        f"({c['recent_pubs'][0]['year']}) {c['recent_pubs'][0]['title']}"
                        if c.get("recent_pubs")
                        else "N/A"
                    ),
                }
                for c in cands
            ]
            st.dataframe(r_rows, use_container_width=True, hide_index=True)

    st.markdown("---")
    st.subheader("📥 Editorial Exports (2 Files)")
    d1, d2 = st.columns(2)
    d1.download_button(
        "📄 Download All-Comments Audit Report (.docx)",
        data=st.session_state["docx_report"],
        file_name="MRJ_All_Comments_Audit_Report.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        use_container_width=True,
    )
    if st.session_state.get("blind_bytes") is not None:
        d2.download_button(
            "🙈 Download Blind Reviewer Copy (.docx)",
            data=st.session_state["blind_bytes"],
            file_name="MRJ_Blind_Reviewer_Copy.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
        )
    else:
        d2.info("Blind Reviewer Copy (.docx) is generated when the uploaded manuscript is a .docx file.")
