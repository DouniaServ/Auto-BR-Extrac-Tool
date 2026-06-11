# extraction_control.py

from __future__ import annotations

import re
import unicodedata
from typing import Dict, List, Optional, Set, Any, Tuple

import fitz  # PyMuPDF
import pandas as pd


# =============================================================================
# Purpose
# =============================================================================
"""
Full Source Traceability / Full PDF Raw Extraction

Quality objective:
- Extract all readable PDF text.
- Do not skip anything.
- Do not classify content as expected / not expected.
- Do not use "Skipped by design".
- Keep all raw PDF content visible by page and reading order.
- Reconcile raw PDF content against the structured Extracted sheet using strict,
  language-neutral matching.
- Keep the output simple: raw text, source group, reconciliation status, and matching reference.

This sheet is a traceability and completeness support layer.

Important principle:
- A wrong reconciliation is worse than keeping an item visible as not reconciled.
- The sheet must remain safe for English, French, Polish, or any other BR language.
- Therefore, matching is strict and does not rely on language-specific keywords.
"""


# =============================================================================
# Status labels
# =============================================================================

MATCHED = "Reconciled with structured pre-draft"
PARTIALLY_MATCHED = "Visible in source - not reconciled"  # kept only for backward compatibility
NOT_MATCHED = "Visible in source - not reconciled"


# =============================================================================
# Output columns expected by app.py
# =============================================================================

CONTROL_COLUMNS = [
    "Page",
    "Source ID",
    "Raw PDF text",
    "Source group",
    "Reconciliation status",
    "Matched Extracted reference",
    "Reviewer note",
]

# Backward-compatible column aliases used internally only
INTERNAL_MATCHED = MATCHED
INTERNAL_NOT_MATCHED = NOT_MATCHED

# =============================================================================
# Regex patterns - language neutral
# =============================================================================
# Unicode words, numbers, document codes, symbols.
TOKEN_PATTERN = re.compile(
    r"(?u)[^\W_]+(?:[./-][^\W_]+)*|≤|≥|±|<|>|%"
)

# Multilingual-ish value pattern. This is not used to classify language;
# it only helps extract critical numeric value tokens as additional raw evidence.
CRITICAL_VALUE_PATTERN = re.compile(
    r"("
    r"\d+(?:[.,]\d+)*\s*(?:kg|g|mg|l|ml|rpm|tr/min|obr/min|hz|min|sec|s|h|hrs|mbar|hpa|°c|c|%|n|mm|cm|bar|g/min|tabs/hr|tablets|comprimés|comprimes|szt)"
    r"|"
    r"\d+(?:[.,]\d+)*\s*(?:-|–|—|to|à|a|do)\s*\d+(?:[.,]\d+)*\s*(?:kg|g|mg|l|ml|rpm|tr/min|obr/min|hz|min|sec|s|h|hrs|mbar|hpa|°c|c|%|n|mm|cm|bar|g/min|tabs/hr|tablets|comprimés|comprimes|szt)?"
    r"|"
    r"[<>≤≥]\s*\d+(?:[.,]\d+)*\s*(?:kg|g|mg|l|ml|rpm|tr/min|obr/min|hz|min|sec|s|h|hrs|mbar|hpa|°c|c|%|n|mm|cm|bar|g/min|tabs/hr|tablets|comprimés|comprimes|szt)?"
    r"|"
    r"\d+(?:[.,]\d+)*\s*±\s*\d+(?:[.,]\d+)*\s*(?:kg|g|mg|l|ml|rpm|tr/min|obr/min|hz|min|sec|s|h|hrs|mbar|hpa|°c|c|%|n|mm|cm|bar|g/min|tabs/hr|tablets|comprimés|comprimes|szt)?"
    r")",
    re.IGNORECASE,
)


# =============================================================================
# Text normalization
# =============================================================================

def strip_accents(text: str) -> str:
    """
    Language-neutral comparison helper:
    température == temperature
    numer == numer
    """
    text = unicodedata.normalize("NFKD", str(text))
    return "".join(ch for ch in text if not unicodedata.combining(ch))


def normalize_text(text: str) -> str:
    if text is None:
        return ""

    text = str(text)

    replacements = {
        "\n": " ",
        "\r": " ",
        "\t": " ",
        "\xa0": " ",
        "": "±",
        "+/-": "±",
        "˚": "°",
        "º": "°",
        "": ">",
        "˃": ">",
        "": "≤",
        "": "8",
        "–": "-",
        "—": "-",
        "−": "-",
        "µ": "u",
        "μ": "u",
        "": "",
        "☐": "",
        "☑": "",
        "☒": "",
        "□": "",
    }

    for old, new in replacements.items():
        text = text.replace(old, new)

    text = strip_accents(text)
    text = re.sub(r"[|•·▪■►]", " ", text)
    text = re.sub(r"\s+", " ", text)

    return text.strip().lower()


def tokenize(text: str) -> List[str]:
    text = normalize_text(text)
    tokens = TOKEN_PATTERN.findall(text)

    cleaned = []
    for token in tokens:
        token = token.strip().lower()

        if not token:
            continue

        # Keep document codes and numeric strings.
        # Remove only isolated non-useful symbols.
        if len(token) == 1 and not token.isdigit() and token not in {
            "%",
            "<",
            ">",
            "≤",
            "≥",
            "±",
        }:
            continue

        cleaned.append(token)

    return cleaned


def token_set(text: str) -> Set[str]:
    return set(tokenize(text))


def missing_tokens_from_sets(raw_tokens: Set[str], extracted_tokens: Set[str]) -> str:
    missing = sorted(raw_tokens - extracted_tokens)
    return ", ".join(missing)


# =============================================================================
# Column helpers
# =============================================================================

def find_first_existing_column(df: pd.DataFrame, candidates: List[str]) -> Optional[str]:
    if not isinstance(df, pd.DataFrame):
        return None

    lower_map = {str(c).strip().lower(): c for c in df.columns}

    for candidate in candidates:
        key = candidate.strip().lower()
        if key in lower_map:
            return lower_map[key]

    return None


def get_data_title_col(df: pd.DataFrame) -> Optional[str]:
    return find_first_existing_column(
        df,
        [
            "Data title",
            "Data Title",
            "Data title ",
            "Data Title ",
        ],
    )


def get_tag_col(df: pd.DataFrame) -> Optional[str]:
    return find_first_existing_column(
        df,
        [
            "Tag",
            "Data tag",
            "Data Tag",
            "Value",
            "Parameter",
        ],
    )


def safe_cell(row: pd.Series, col: Optional[str]) -> str:
    if not col:
        return ""

    if col not in row.index:
        return ""

    value = row.get(col, "")

    if pd.isna(value):
        return ""

    return str(value).strip()


# =============================================================================
# Raw PDF extraction
# =============================================================================

def extract_all_raw_pdf_items(pdf_path: str) -> pd.DataFrame:
    """
    Extract all readable PDF text without applying business skip logic.

    Kept:
    - TEXT_LINE: PyMuPDF text extraction line-by-line.
    - VISUAL_WORD_ROW: visual reconstruction by y-position.
    - CRITICAL_VALUE_TOKEN: numeric values with units/ranges.

    Nothing is classified as expected/not expected.
    Nothing is skipped intentionally, except exact duplicates from the same page/source.
    """

    rows = []
    seen = set()

    def add_row(page_number: int, raw_order: int, source_type: str, text: str):
        raw = str(text).strip()

        if not raw:
            return

        norm_key = re.sub(r"\s+", " ", raw).strip().lower()

        key = (
            int(page_number),
            str(source_type),
            norm_key,
        )

        if key in seen:
            return

        seen.add(key)

        rows.append({
            "Page number": int(page_number),
            "Raw order": int(raw_order),
            "Raw source type": str(source_type),
            "Raw PDF text": raw,
        })

    doc = fitz.open(pdf_path)

    try:
        for page_index in range(len(doc)):
            page = doc[page_index]
            page_number = page_index + 1
            raw_order = 0

            # -----------------------------------------------------------------
            # 1. TEXT_LINE extraction
            # -----------------------------------------------------------------
            text = page.get_text("text") or ""

            for line in text.splitlines():
                raw_order += 1
                add_row(page_number, raw_order, "TEXT_LINE", line)

            # -----------------------------------------------------------------
            # 2. VISUAL_WORD_ROW extraction
            # -----------------------------------------------------------------
            words = page.get_text("words") or []
            words = sorted(
                words,
                key=lambda w: (round(w[1], 1), round(w[0], 1)),
            )

            visual_rows = []
            y_tolerance = 3.0

            for word_info in words:
                x0, y0, x1, y1, word = word_info[:5]
                word = str(word).strip()

                if not word:
                    continue

                placed = False

                for group in visual_rows:
                    if abs(group["y"] - y0) <= y_tolerance:
                        group["words"].append((x0, word))
                        placed = True
                        break

                if not placed:
                    visual_rows.append({
                        "y": y0,
                        "words": [(x0, word)],
                    })

            visual_page_texts = []

            for group in visual_rows:
                group_words = sorted(group["words"], key=lambda x: x[0])
                visual_text = " ".join([w for _, w in group_words]).strip()

                if visual_text:
                    visual_page_texts.append(visual_text)

                raw_order += 1
                add_row(page_number, raw_order, "VISUAL_WORD_ROW", visual_text)

            # -----------------------------------------------------------------
            # 3. CRITICAL_VALUE_TOKEN extraction
            # -----------------------------------------------------------------
            value_source_text = " ".join(visual_page_texts)

            for match in CRITICAL_VALUE_PATTERN.finditer(value_source_text):
                value = str(match.group(0)).strip()

                if not value:
                    continue

                raw_order += 1
                add_row(page_number, raw_order, "CRITICAL_VALUE_TOKEN", value)

    finally:
        doc.close()

    raw_df = pd.DataFrame(rows)

    if raw_df.empty:
        return pd.DataFrame(columns=[
            "Page number",
            "Raw order",
            "Raw source type",
            "Raw PDF text",
        ])

    raw_df = raw_df.sort_values(
        by=["Page number", "Raw order", "Raw source type"],
        kind="stable",
    ).reset_index(drop=True)

    return raw_df


# Backward-compatible function name if app.py still calls the old one.
def extract_raw_pdf_items(pdf_path: str) -> pd.DataFrame:
    return extract_all_raw_pdf_items(pdf_path)


# =============================================================================
# Extracted sheet index
# =============================================================================

def get_extractable_columns(df: pd.DataFrame) -> List[str]:
    """
    Columns used for source traceability matching.

    To avoid false positives across English/French/Polish:
    - Do not match on Process step / Sub process step because product names
      like "Daflon" create false matches.
    - Do not match on Tag / Data tag because values such as TEXT or INSTRUCTION
      are output categories, not raw PDF source text.
    - Keep source-like columns such as Data title, Value, Théorique, Réel, Visa, etc.
    """
    ignored = {
        "PDF name",
        "Data number",
        "Page number",
        "Paragraph number",
        "Process step",
        "Sub process step",
        "Tag",
        "Data tag",
        "Data Tag",
        "Reviewed & OK",
        "Reviewer",
        "Reviewed at",
        "Reviewer note",
    }

    return [c for c in df.columns if str(c).strip() not in ignored]


def build_full_row_text(row: pd.Series, usable_cols: List[str]) -> str:
    values = []

    for col in usable_cols:
        value = row.get(col, "")

        if pd.isna(value):
            continue

        value = str(value).strip()

        if value:
            values.append(value)

    return " | ".join(values).strip()


def build_extracted_field_index(extracted_df: pd.DataFrame) -> pd.DataFrame:
    """
    Build index from selected source-like fields in Extracted sheet.
    """

    if not isinstance(extracted_df, pd.DataFrame) or extracted_df.empty:
        return pd.DataFrame(columns=[
            "Candidate id",
            "Extracted row",
            "Page number",
            "Field",
            "Value",
            "Norm value",
            "Token set",
            "Data title",
            "Tag",
            "Full row",
        ])

    rows = []
    extractable_cols = get_extractable_columns(extracted_df)
    data_title_col = get_data_title_col(extracted_df)
    tag_col = get_tag_col(extracted_df)

    candidate_id = 0

    for idx, row in extracted_df.iterrows():
        page_number = None

        if "Page number" in extracted_df.columns:
            page_number = pd.to_numeric(row.get("Page number"), errors="coerce")
            page_number = None if pd.isna(page_number) else int(page_number)

        data_title = safe_cell(row, data_title_col)
        tag = safe_cell(row, tag_col)
        full_row = build_full_row_text(row, extractable_cols)

        for col in extractable_cols:
            value = row.get(col, "")

            if pd.isna(value):
                continue

            value = str(value).strip()

            if not value:
                continue

            tokens = token_set(value)

            if not tokens:
                continue

            rows.append({
                "Candidate id": int(candidate_id),
                "Extracted row": int(idx),
                "Page number": page_number,
                "Field": str(col),
                "Value": value,
                "Norm value": normalize_text(value),
                "Token set": tokens,
                "Data title": data_title,
                "Tag": tag,
                "Full row": full_row,
            })

            candidate_id += 1

    return pd.DataFrame(rows)


def build_extracted_row_index(extracted_df: pd.DataFrame) -> pd.DataFrame:
    """
    Build full-row index from selected source-like fields only.
    """

    if not isinstance(extracted_df, pd.DataFrame) or extracted_df.empty:
        return pd.DataFrame(columns=[
            "Candidate id",
            "Extracted row",
            "Page number",
            "Field",
            "Value",
            "Norm value",
            "Token set",
            "Data title",
            "Tag",
            "Full row",
        ])

    usable_cols = get_extractable_columns(extracted_df)
    data_title_col = get_data_title_col(extracted_df)
    tag_col = get_tag_col(extracted_df)

    rows = []
    candidate_id = 0

    for idx, row in extracted_df.iterrows():
        page_number = None

        if "Page number" in extracted_df.columns:
            page_number = pd.to_numeric(row.get("Page number"), errors="coerce")
            page_number = None if pd.isna(page_number) else int(page_number)

        data_title = safe_cell(row, data_title_col)
        tag = safe_cell(row, tag_col)
        full_row = build_full_row_text(row, usable_cols)

        if not full_row:
            continue

        tokens = token_set(full_row)

        if not tokens:
            continue

        rows.append({
            "Candidate id": int(candidate_id),
            "Extracted row": int(idx),
            "Page number": page_number,
            "Field": "__FULL_EXTRACTED_ROW__",
            "Value": full_row,
            "Norm value": normalize_text(full_row),
            "Token set": tokens,
            "Data title": data_title,
            "Tag": tag,
            "Full row": full_row,
        })

        candidate_id += 1

    return pd.DataFrame(rows)


def build_combined_extracted_index(extracted_df: pd.DataFrame) -> pd.DataFrame:
    field_index = build_extracted_field_index(extracted_df)
    row_index = build_extracted_row_index(extracted_df)

    combined = pd.concat(
        [field_index, row_index],
        ignore_index=True,
    )

    if combined.empty:
        return combined

    combined["Candidate id"] = range(len(combined))

    return combined


def build_fast_match_index(extracted_index: pd.DataFrame) -> Dict[str, Any]:
    index: Dict[str, Any] = {
        "all": {
            "records": {},
            "token_to_ids": {},
        },
        "pages": {},
    }

    if extracted_index is None or extracted_index.empty:
        return index

    records = extracted_index.to_dict("records")

    for rec in records:
        candidate_id = int(rec["Candidate id"])
        page_number = rec.get("Page number", None)
        tokens = rec.get("Token set", set())

        if not isinstance(tokens, set):
            tokens = token_set(rec.get("Value", ""))

        index["all"]["records"][candidate_id] = rec

        for token in tokens:
            index["all"]["token_to_ids"].setdefault(token, set()).add(candidate_id)

        if page_number is not None and not pd.isna(page_number):
            page_number = int(page_number)

            if page_number not in index["pages"]:
                index["pages"][page_number] = {
                    "records": {},
                    "token_to_ids": {},
                }

            index["pages"][page_number]["records"][candidate_id] = rec

            for token in tokens:
                index["pages"][page_number]["token_to_ids"].setdefault(token, set()).add(candidate_id)

    return index


# =============================================================================
# Strict language-neutral matching
# =============================================================================

def blank_match() -> Dict[str, Any]:
    return {
        "Extracted row": "",
        "Field": "",
        "Value": "",
        "Score": 0.0,
        "Coverage": 0.0,
        "Data title": "",
        "Tag": "",
        "Full row": "",
    }


UNIT_TOKENS = {
    "kg", "g", "mg", "l", "ml", "rpm", "tr/min", "obr/min",
    "hz", "min", "sec", "s", "h", "hrs", "mbar", "hpa",
    "c", "°c", "n", "mm", "cm", "bar", "%", "tablets",
    "comprimes", "comprimés", "szt",
}


def is_candidate_token(token: str) -> bool:
    """
    Language-neutral token filter for candidate search.
    Avoids candidates based only on units, single digits, or very short tokens.
    """
    t = str(token).strip().lower()

    if not t:
        return False

    if t in UNIT_TOKENS:
        return False

    # Do not use a single character token to select candidates.
    if len(t) == 1:
        return False

    # Avoid using plain 1-2 digit numbers as candidate drivers.
    if t.isdigit() and len(t) <= 2:
        return False

    return True


def normalized_equal(a: str, b: str) -> bool:
    a_norm = normalize_text(a)
    b_norm = normalize_text(b)
    return bool(a_norm) and a_norm == b_norm


def boundary_contains(haystack: str, needle: str) -> bool:
    """
    Phrase containment with non-word boundaries.
    Prevents matching '4' inside '604288-4'.
    """
    h = normalize_text(haystack)
    n = normalize_text(needle)

    if not h or not n:
        return False

    pattern = r"(?<![\w./-])" + re.escape(n) + r"(?![\w./-])"
    return re.search(pattern, h, flags=re.UNICODE) is not None


def token_coverage(source: str, target: str) -> float:
    source_tokens = token_set(source)
    target_tokens = token_set(target)

    if not source_tokens:
        return 0.0

    return len(source_tokens & target_tokens) / len(source_tokens)


def strict_phrase_match(raw_text: str, candidate_text: str) -> Tuple[bool, float]:
    """
    Strict, language-neutral matching.

    Accept only:
    1. Exact normalized match.
    2. Raw text is contained in candidate text if it is not too short and covers
       a large part of the candidate.
    3. Candidate text is contained in raw text if it is meaningful enough and
       covers a large part of the raw text.

    Reject examples:
    - BATCH RECORD -> long instruction containing Batch Record
    - Product Name -> Stage 04 Coating
    - DAFLON 1000 MG -> daflon
    - 806.8 kg / 622,500 Tablets -> 7.470 kg
    - 604288-4 -> 4
    """

    raw_norm = normalize_text(raw_text)
    cand_norm = normalize_text(candidate_text)

    if not raw_norm or not cand_norm:
        return False, 0.0

    if raw_norm == cand_norm:
        return True, 1.0

    raw_len = len(raw_norm)
    cand_len = len(cand_norm)

    raw_tokens = token_set(raw_text)
    cand_tokens = token_set(candidate_text)

    # Reject very weak comparisons.
    if raw_len < 6 or cand_len < 6:
        return False, 0.0

    # Raw phrase inside candidate.
    # Needs high ratio to avoid "BATCH RECORD" matching a long instruction.
    if boundary_contains(cand_norm, raw_norm):
        ratio = raw_len / max(cand_len, 1)
        cov = token_coverage(raw_text, candidate_text)

        if ratio >= 0.65 and cov >= 0.90:
            return True, min(0.95, cov)

    # Candidate phrase inside raw.
    # Allows source row "Mixer speed 60 ± 2 RPM" to match extracted value
    # "Mixer speed" or "60 ± 2 RPM" only if it is meaningful enough.
    if boundary_contains(raw_norm, cand_norm):
        ratio = cand_len / max(raw_len, 1)
        cov = token_coverage(candidate_text, raw_text)

        # Critical value candidates can be shorter but must be exact phrase.
        if CRITICAL_VALUE_PATTERN.search(candidate_text):
            if ratio >= 0.30 and cov >= 0.90:
                return True, min(0.95, cov)

        # General text candidate must cover enough of raw line.
        if ratio >= 0.50 and cov >= 0.90:
            return True, min(0.90, cov)

    return False, 0.0


def safe_traceability_match(raw_text: str, extracted_value: str, extracted_full_row: str) -> Tuple[bool, float]:
    """
    Match raw text against extracted field and full row.
    Returns (accepted, score).
    """

    ok_value, score_value = strict_phrase_match(raw_text, extracted_value)

    if ok_value:
        return True, score_value

    ok_row, score_row = strict_phrase_match(raw_text, extracted_full_row)

    if ok_row:
        return True, score_row

    return False, 0.0


def best_match(
    raw_text: str,
    raw_page: int,
    fast_index: Optional[Dict[str, Any]] = None,
) -> Dict[str, Any]:
    """
    Strict Quality traceability matching.

    No business classification.
    No skipped logic.
    No weak partial matching.

    Better result:
    - false positives become Not covered
    - exact source evidence remains reconciled with the structured pre-draft
    """

    raw_tokens = token_set(raw_text)

    if not raw_tokens:
        return blank_match()

    if fast_index is None:
        return blank_match()

    # Prefer same page matching.
    bucket = fast_index.get("pages", {}).get(int(raw_page))

    if not bucket:
        bucket = fast_index.get("all", {})

    records = bucket.get("records", {})
    token_to_ids = bucket.get("token_to_ids", {})

    candidate_ids = set()

    for token in raw_tokens:
        if not is_candidate_token(token):
            continue

        candidate_ids.update(token_to_ids.get(token, set()))

    # If no strong candidate token, do not force a match.
    if not candidate_ids:
        return blank_match()

    best = blank_match()

    for candidate_id in candidate_ids:
        candidate = records.get(candidate_id)

        if not candidate:
            continue

        extracted_value = str(candidate.get("Value", ""))
        extracted_full_row = str(candidate.get("Full row", ""))
        field_name = str(candidate.get("Field", ""))

        accepted, score = safe_traceability_match(
            raw_text=raw_text,
            extracted_value=extracted_value,
            extracted_full_row=extracted_full_row,
        )

        if not accepted:
            continue

        # Coverage of raw text by the accepted candidate evidence.
        coverage_value = token_coverage(raw_text, extracted_value)
        coverage_row = token_coverage(raw_text, extracted_full_row)
        coverage = max(coverage_value, coverage_row)

        if score > float(best["Score"]):
            best = {
                "Extracted row": candidate.get("Extracted row", ""),
                "Field": field_name,
                "Value": extracted_value,
                "Score": float(score),
                "Coverage": float(coverage),
                "Data title": str(candidate.get("Data title", "")),
                "Tag": str(candidate.get("Tag", "")),
                "Full row": extracted_full_row,
            }

    return best


def determine_match_status(
    raw_source_type: str,
    score: float,
    coverage: float,
    threshold_matched: float,
    threshold_partial: float,
) -> str:
    """
    Strict status logic.
    Only 'Reconciled with structured pre-draft' or 'Visible in source - not reconciled'.

    This is language-neutral and safer for Quality than broad partial matching.
    """

    if score >= 0.90 and coverage >= 0.50:
        return MATCHED

    return NOT_MATCHED



# =============================================================================
# Quality comfort classification layer
# =============================================================================

DATE_PATTERN = re.compile(
    r"\b(\d{1,2}[/-]\d{1,2}[/-]\d{2,4}|\d{1,2}\s+[a-zA-Zéûîôàèùç]+\s+\d{4}|\d{4}-\d{2}-\d{2})\b",
    re.IGNORECASE,
)

EMAIL_PATTERN = re.compile(r"\b[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}\b", re.IGNORECASE)
IP_PATTERN = re.compile(r"\b\d{1,3}(?:\.\d{1,3}){3}\b")
PAGE_PATTERN = re.compile(r"\b(page|strona|str\.?|pagina|página)\s+\d+\s*(of|sur|z|/)\s*\d+\b", re.IGNORECASE)
DOCUSIGN_PATTERN = re.compile(r"\b(docusign|envelope|certificate of completion|signature adoption|security level|ip address|signed:|viewed:|sent:|completed)\b", re.IGNORECASE)
SIGNATURE_PATTERN = re.compile(r"\b(signature|signed|prepared by|reviewed by|approved by|author|qa operations|trainee|trained by|conclusion|podpis|zatwierdz|approuv|visa)\b", re.IGNORECASE)
HEADER_LABEL_PATTERN = re.compile(r"\b(product name|item number|document number|batch number|batch record|nom du produit|numero|numéro|dokument|numer|produktu)\b", re.IGNORECASE)


def normalized_for_repetition(text: str) -> str:
    return normalize_text(text)


def build_repetition_counts(raw_df: pd.DataFrame) -> Dict[str, int]:
    if raw_df is None or raw_df.empty or "Raw PDF text" not in raw_df.columns:
        return {}

    counts: Dict[str, int] = {}
    # Count a text once per page so repeated headers are detected reliably.
    pairs = set()
    for _, r in raw_df.iterrows():
        page = int(r.get("Page number", 0) or 0)
        norm = normalized_for_repetition(str(r.get("Raw PDF text", "")))
        if not norm:
            continue
        pairs.add((page, norm))

    for _, norm in pairs:
        counts[norm] = counts.get(norm, 0) + 1

    return counts


def classify_source_content(
    raw_text: str,
    raw_source_type: str,
    raw_page: int,
    raw_order: int,
    repeat_count: int,
    total_pages: int,
) -> Tuple[str, str, str]:
    """
    Deterministic, auditable classification for Quality comfort.

    Important:
    - This does not hide or skip content.
    - It only separates visible content into likely categories.
    - Anything that may contain process/value information remains reviewer-assessment.
    """
    raw = str(raw_text or "").strip()
    norm = normalize_text(raw)
    source = str(raw_source_type or "").upper()
    tokens = tokenize(raw)
    token_count = len(tokens)
    has_number = any(t.isdigit() or re.search(r"\d", t) for t in tokens)
    has_critical = bool(CRITICAL_VALUE_PATTERN.search(raw))

    # 1. Digital signing / certificate / audit metadata.
    if DOCUSIGN_PATTERN.search(raw) or EMAIL_PATTERN.search(raw) or IP_PATTERN.search(raw):
        return (
            "Document / e-signature metadata",
            "Visible - likely out of structured scope",
            "DocuSign/certificate/audit metadata is visible for traceability but is normally outside the structured BR simplification output.",
        )

    # 2. Page metadata.
    if PAGE_PATTERN.search(raw) or re.fullmatch(r"page\s*\d+\s*(of|sur|z|/)\s*\d+", norm or ""):
        return (
            "Page metadata",
            "Visible - likely out of structured scope",
            "Page numbering is visible but normally not transferred as structured process data.",
        )

    # 3. Repeated document header/footer content.
    # Use repetition across pages to stay language-neutral.
    if repeat_count >= 3 and (raw_order <= 12 or HEADER_LABEL_PATTERN.search(raw)):
        return (
            "Repeated header/footer",
            "Visible - likely out of structured scope",
            "Repeated document header/footer content is visible; normally not a missing structured extraction item.",
        )

    # 4. Signature / approval / training blocks.
    if SIGNATURE_PATTERN.search(raw):
        return (
            "Signature / approval / training block",
            "Visible - likely out of structured scope",
            "Signature, approval or training information is visible; normally reviewed as document metadata, not as structured simplification content.",
        )

    # 5. Critical values / numeric process data: keep as reviewer assessment if not covered.
    if source == "CRITICAL_VALUE_TOKEN" or has_critical:
        return (
            "Potential process value / parameter",
            "Visible - requires reviewer assessment",
            "Numeric value or unit detected. If it is in scope, reviewer should confirm it is present in the structured output.",
        )

    # 6. Short labels/structural table headings.
    # These are visible but normally not each a separate structured item.
    if token_count <= 4 and not has_critical:
        if not has_number or HEADER_LABEL_PATTERN.search(raw):
            return (
                "Table label / structural text",
                "Visible - likely out of structured scope",
                "Short label or table heading is visible; normally it provides context rather than a standalone structured value.",
            )

    # 7. Default: content may be in scope and should be assessed if not covered.
    return (
        "Main content / potential BR data",
        "Visible - requires reviewer assessment",
        "Source content is visible but no strict structured match was found. Reviewer confirms whether it is in scope.",
    )


def reviewer_priority_for(coverage_status: str, category: str, raw_text: str) -> str:
    if coverage_status == MATCHED:
        return "Low"

    if coverage_status == "Visible - likely out of structured scope":
        return "Low"

    if CRITICAL_VALUE_PATTERN.search(str(raw_text or "")):
        return "High"

    if "Main content" in str(category) or "Potential process value" in str(category):
        return "Medium"

    return "Low"


def reviewer_action_for(coverage_status: str, category: str) -> str:
    if coverage_status == MATCHED:
        return "No action unless reviewer wants to spot-check."

    if coverage_status == "Visible - likely out of structured scope":
        return "No extraction correction expected; confirm only if this item is considered in scope."

    return "Review against the PDF and add/correct the Extracted sheet if this item is in scope."



def traceability_group_and_status(
    match_status: str,
    source_category: str,
) -> Tuple[str, str]:
    """
    Simplified Quality-facing logic.

    The goal is not to create a risk table. The goal is to show:
    - what raw source text was read
    - whether it was reconciled with the structured pre-draft
    - whether non-reconciled content is only document metadata/context
    """
    if match_status == MATCHED:
        return "Source data", "Reconciled with structured pre-draft"

    cat = str(source_category or "").lower()
    metadata_terms = [
        "metadata",
        "header",
        "footer",
        "signature",
        "approval",
        "training",
        "table label",
        "structural",
        "page",
        "docusign",
        "certificate",
    ]

    if any(term in cat for term in metadata_terms):
        return "Document metadata / context", "Document metadata / context"

    return "Source data", "Visible in source - not reconciled"


def format_matched_reference(match: Dict[str, Any]) -> str:
    """
    Compact reference for the reviewer.
    Keep only useful evidence; no technical score/token columns.
    """
    row = match.get("Extracted row", "")
    field = str(match.get("Field", "") or "").strip()
    value = str(match.get("Value", "") or "").strip()
    title = str(match.get("Data title", "") or "").strip()

    parts = []

    try:
        if row != "" and row is not None:
            parts.append(f"Extracted row {int(row) + 2}")
    except Exception:
        pass

    if field and field != "__FULL_EXTRACTED_ROW__":
        parts.append(field)

    if title:
        parts.append(title)

    if value:
        short_value = value if len(value) <= 120 else value[:117] + "..."
        parts.append(short_value)

    return " | ".join(parts)

# =============================================================================
# Main function used by your app
# =============================================================================

def build_full_pdf_raw_extraction_control(
    pdf_path: str,
    extracted_df: pd.DataFrame,
    threshold_matched: float = 0.90,
    threshold_partial: float = 0.80,
    include_summary: bool = False,
    **kwargs,
) -> pd.DataFrame:
    """
    Build a simple Full Source Traceability sheet.

    One row = one readable PDF raw item.

    Quality-facing design:
    - The Extracted sheet remains the main pre-draft review sheet.
    - This sheet is a traceability support layer, not a second mandatory review.
    - It keeps all readable source text visible.
    - It separates obvious document metadata/context from source data.
    - It avoids risk wording, priority scoring, and technical matching columns.
    """

    if not isinstance(extracted_df, pd.DataFrame):
        extracted_df = pd.DataFrame()

    raw_df = extract_all_raw_pdf_items(pdf_path)

    combined_index = build_combined_extracted_index(extracted_df)
    fast_index = build_fast_match_index(combined_index)

    repeat_counts = build_repetition_counts(raw_df)
    total_pages = int(raw_df["Page number"].max()) if not raw_df.empty and "Page number" in raw_df.columns else 0

    control_rows = []

    for _, raw_row in raw_df.iterrows():
        raw_page = int(raw_row["Page number"])
        raw_order = int(raw_row["Raw order"])
        raw_source_type = str(raw_row.get("Raw source type", "TEXT_LINE"))
        raw_text = str(raw_row.get("Raw PDF text", "")).strip()

        match = best_match(
            raw_text=raw_text,
            raw_page=raw_page,
            fast_index=fast_index,
        )

        score = float(match.get("Score", 0.0))
        coverage = float(match.get("Coverage", 0.0))

        match_status = determine_match_status(
            raw_source_type=raw_source_type,
            score=score,
            coverage=coverage,
            threshold_matched=threshold_matched,
            threshold_partial=threshold_partial,
        )

        repeat_key = normalized_for_repetition(raw_text)
        repeat_count = int(repeat_counts.get(repeat_key, 0))

        source_category, _default_status, _rationale = classify_source_content(
            raw_text=raw_text,
            raw_source_type=raw_source_type,
            raw_page=raw_page,
            raw_order=raw_order,
            repeat_count=repeat_count,
            total_pages=total_pages,
        )

        source_group, reconciliation_status = traceability_group_and_status(
            match_status=match_status,
            source_category=source_category,
        )

        if match_status != MATCHED:
            match = blank_match()

        control_rows.append({
            "Page": raw_page,
            "Source ID": f"P{raw_page:03d}-{raw_order:04d}",
            "Raw PDF text": raw_text,
            "Source group": source_group,
            "Reconciliation status": reconciliation_status,
            "Matched Extracted reference": format_matched_reference(match) if match_status == MATCHED else "",
            "Reviewer note": "",
        })

    control_df = pd.DataFrame(control_rows, columns=CONTROL_COLUMNS)

    return control_df


# =============================================================================
# Backward-compatible function name
# =============================================================================

def build_strong_raw_order_control(
    pdf_path: str,
    extracted_df: pd.DataFrame,
    threshold_extracted: float = 0.90,
    threshold_partial: float = 0.80,
    include_summary: bool = False,
    **kwargs,
) -> pd.DataFrame:
    """
    Backward-compatible wrapper.

    If app.py already calls build_strong_raw_order_control(),
    it now generates the language-neutral Full Source Coverage Matrix.
    """

    return build_full_pdf_raw_extraction_control(
        pdf_path=pdf_path,
        extracted_df=extracted_df,
        threshold_matched=threshold_extracted,
        threshold_partial=threshold_partial,
        include_summary=include_summary,
        **kwargs,
    )
