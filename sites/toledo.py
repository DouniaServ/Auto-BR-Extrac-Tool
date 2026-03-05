


import os
import re
import tempfile
import unicodedata
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd
import pdfplumber


# ============================================================
# UI OUTPUT COLUMNS
# ============================================================
COLUMNS = [
    "Unnamed: 0",
    "Data number",
    "Page number",
    "Paragraph number",
    "Process step",
    "Sub process step",
    "Data title",
    "Tag",
]

# ============================================================
# Allowed Tag data values (HARD CONSTRAINT)
# ============================================================
ALLOWED_TAG_DATA = {
    "TEXT",
    "INSTRUCTION",
    "DATE",
    "FORMULA",
    "IPC_HEADER_STRUCT",
    "VPR_STAGE_NAME",
    "VPR_EQUIPMENT",
    "VPR_EQUIPMENT_TEMPLATE",
    "PREP_CODE",
    "MATERIAL_NAME",
    "ITEM_CODE",
    "REQUIREMENTS",
    "QUANTITY_REQUIRED",
    "PROCESS_INSTRUCTION",
    "CHECKBOX_CHOICE",
    "YES_CHECKBOX",
    "NO_CHECKBOX",
    "NA_CHECKBOX",
    "YES_NA_CHECKBOX",
    "YES_NO_CHECKBOX",
    "NO_NA_CHECKBOX",
    "YES_NO_NA_CHECKBOX",
}

def tag_data_safe(tag: str) -> str:
    t = (tag or "").strip().upper()
    return t if t in ALLOWED_TAG_DATA else "TEXT"

# ============================================================
# Tags in Spanish (EXTRACTION OUTPUT)
# ============================================================
TAG_LEFT_ES = "Instrucciones"
TAG_RIGHT_ES = "Registro de parámetros"
TAG_TEXT_LIBRE_ES = "Texto libre"

# UI mapping (adapter only; extraction logic unchanged)
TAG_ES_TO_UI = {
    TAG_LEFT_ES: "INSTRUCTION",
    TAG_RIGHT_ES: "TEXT",
    TAG_TEXT_LIBRE_ES: "TEXT",
}

# ============================================================
# Helpers / Normalization
# ============================================================
_re_multispace = re.compile(r"\s+")
DOTLEADER_RX = re.compile(r"\.{6,}|…{2,}")
ZERO_WIDTH_RE = re.compile(r"[\u200B-\u200F\u202A-\u202E\u2060\uFEFF]")
NBSP_RE = re.compile(r"[\u00A0\u202F]")

HEADER_NOISE_RX = re.compile(
    r"^(SERVIER|P[ÁA]GINA|MOF\.|REVISADO|COPIA CONTROLADA|EDICI[ÓO]N)\b"
    r"|^Page\s+\d+\s+(?:sur|of)\s+\d+",
    re.IGNORECASE,
)

HEADER_CONTAINS = [
    "marcha operación", "marcha operacion",
    "hecho",
    "observaciones",
    "firma", "fecha", "hora",
]

NOISE_TOKENS = {"", "-", "—", "↓", "→", "x", "✓", "☒", "☐", "si", "sí", "no", "n/a", "na"}

SUBPROCESS_RX = re.compile(r"^\s*(\d+)\s*[-–]\s*(.+\S)\s*$")

PRODUCT_RXES = [
    re.compile(r"\b(9490\s+PERINDARG\s+6S\s+TO)\b", re.IGNORECASE),
    re.compile(r"\b(PERINDARG\s+6S\s+TO)\b", re.IGNORECASE),
    re.compile(r"\b(9490\s+PERINDARG(?:\s+\S+){0,6})\b", re.IGNORECASE),
]


def norm(s: str) -> str:
    s = s or ""
    s = unicodedata.normalize("NFKC", s)
    s = NBSP_RE.sub(" ", s)
    s = ZERO_WIDTH_RE.sub("", s)
    return re.sub(r"\s+", " ", s).strip()

def flat(s: str) -> str:
    return _re_multispace.sub(" ", (s or "").replace("\u00a0", " ")).strip()

def is_header_noise(line: str) -> bool:
    return bool(HEADER_NOISE_RX.match(flat(line)))

def is_ignorable_header(text: str) -> bool:
    t = flat(text).lower().strip(" :;-")
    return any(h in t for h in HEADER_CONTAINS)

def is_noise_token(text: str) -> bool:
    t = flat(text).lower()
    if t in NOISE_TOKENS:
        return True
    return len(t) < 2

def tokenize_simple(text: str) -> List[str]:
    t = flat(text)
    return [w for w in re.split(r"\s+", t) if w] if t else []

def is_meaningful_text(text: str) -> bool:
    t = flat(text)
    if not t:
        return False
    if is_header_noise(t) or is_ignorable_header(t) or is_noise_token(t):
        return False
    if sum(ch.isalpha() for ch in t) < 5:
        return False
    if len(t) < 10:
        return False
    return True


# ============================================================
# Process step / Sub process step
# ============================================================
def extract_product_name(lines: List[str], fallback: str = "--") -> str:
    blob = " ".join(lines)
    for rx in PRODUCT_RXES:
        m = rx.search(blob)
        if m:
            return flat(m.group(1)).upper()

    m2 = re.search(r"FABRICACI[ÓO]N\s+DE\s*[:\-]?\s*(.+)$", blob, re.IGNORECASE)
    if m2:
        val = flat(m2.group(1))
        if 3 <= len(val) <= 80:
            return val.upper()

    return fallback

def extract_subprocess_step(lines: List[str]) -> Optional[str]:
    for l in lines:
        t = flat(l)
        if not t or is_header_noise(t):
            continue
        m = SUBPROCESS_RX.match(t)
        if not m:
            continue
        n = m.group(1)
        title = m.group(2).strip()
        if sum(ch.isalpha() for ch in title) < 5:
            continue
        return f"{n}- {title}"
    return None


# ============================================================
# Split X (left/right)
# ============================================================
def compute_split_x_from_table(table_obj) -> Optional[float]:
    try:
        x0, _, _, _ = table_obj.bbox
        cols = list(table_obj.columns)
        interior_lefts = [c.bbox[0] for c in cols if c.bbox[0] > x0 + 5]
        return min(interior_lefts) if interior_lefts else None
    except Exception:
        return None

def find_header_x_positions(page) -> Dict[str, float]:
    words = page.extract_words(keep_blank_chars=False, use_text_flow=True) or []
    best: Dict[str, float] = {}
    for w in words:
        txt = flat(w.get("text", "")).lower()
        if txt in {"marcha", "hecho", "observaciones"}:
            best[txt] = min(best.get(txt, w["x0"]), w["x0"])
    return best

def compute_split_x_from_headers(page) -> Optional[float]:
    pos = find_header_x_positions(page)
    left = pos.get("marcha")
    right = pos.get("hecho") or pos.get("observaciones")
    if left is None or right is None or right <= left:
        return None
    return (left + right) / 2.0

def estimate_text_x0_in_row(page, row_bbox: Tuple[float, float, float, float], text: str) -> Optional[float]:
    x0, top, x1, bottom = row_bbox
    words = page.extract_words(keep_blank_chars=False, use_text_flow=True, extra_attrs=[]) or []
    tol = 2.0
    row_words = [
        w for w in words
        if (w["top"] >= top - tol and w["bottom"] <= bottom + tol and w["x0"] >= x0 - tol and w["x1"] <= x1 + tol)
    ]
    if not row_words:
        return None
    row_words_sorted = sorted(row_words, key=lambda w: w["x0"])
    row_tokens = [flat(w["text"]) for w in row_words_sorted]
    toks = tokenize_simple(text)
    if not toks:
        return None

    if len(toks) >= 2:
        for i in range(len(row_tokens) - 1):
            if row_tokens[i] == toks[0] and row_tokens[i + 1] == toks[1]:
                return row_words_sorted[i]["x0"]

    for i in range(len(row_tokens)):
        if row_tokens[i] == toks[0]:
            return row_words_sorted[i]["x0"]

    return None

def assign_tag_es_from_x(x: Optional[float], split_x: Optional[float], fallback_col_i: int) -> str:
    if split_x is None or x is None:
        return TAG_LEFT_ES if fallback_col_i == 0 else TAG_RIGHT_ES
    return TAG_LEFT_ES if x < split_x else TAG_RIGHT_ES


# ============================================================
# Table cell selection
# ============================================================
def pick_titles_with_col(row: List[Optional[str]]) -> List[Tuple[str, int]]:
    out: List[Tuple[str, int]] = []
    for col_i, cell in enumerate(row or []):
        c = norm(cell or "")
        if not c:
            continue
        if DOTLEADER_RX.search(c):
            continue
        if is_header_noise(c) or is_ignorable_header(c) or is_noise_token(c):
            continue
        if len(flat(c)) < 3:
            continue
        out.append((c, col_i))

    seen = set()
    uniq: List[Tuple[str, int]] = []
    for text, col_i in out:
        k = text.strip().lower()
        if k not in seen:
            uniq.append((text, col_i))
            seen.add(k)
    return uniq


# ============================================================
# Fallback: words -> rows
# ============================================================
def cluster_words_into_rows(words: List[Dict[str, Any]], y_tol: float = 3.0) -> List[List[Dict[str, Any]]]:
    if not words:
        return []
    words = sorted(words, key=lambda w: (w["top"], w["x0"]))
    rows: List[List[Dict[str, Any]]] = []
    cur: List[Dict[str, Any]] = [words[0]]
    cur_y = words[0]["top"]
    for w in words[1:]:
        if abs(w["top"] - cur_y) <= y_tol:
            cur.append(w)
        else:
            rows.append(sorted(cur, key=lambda t: t["x0"]))
            cur = [w]
            cur_y = w["top"]
    rows.append(sorted(cur, key=lambda t: t["x0"]))
    return rows

def row_text(row_words: List[Dict[str, Any]]) -> str:
    return flat(" ".join(w["text"] for w in row_words))

def extract_rows_from_words(page, top_margin: float = 60.0, bottom_margin: float = 40.0) -> List[Tuple[str, float]]:
    words = page.extract_words(keep_blank_chars=False, use_text_flow=True) or []
    rows = cluster_words_into_rows(words, y_tol=3.0)
    out: List[Tuple[str, float]] = []
    H = getattr(page, "height", None)

    for rw in rows:
        if not rw:
            continue
        top = min(w["top"] for w in rw)
        bottom = max(w["bottom"] for w in rw)
        if top < top_margin:
            continue
        if H is not None and bottom > (H - bottom_margin):
            continue

        txt = row_text(rw)
        if not txt:
            continue
        if is_header_noise(txt) or is_ignorable_header(txt):
            continue

        row_x0 = min(w["x0"] for w in rw)
        out.append((txt, row_x0))
    return out


# ============================================================
# Data model
# ============================================================
@dataclass
class DataRow:
    data_number: int
    page_number: int
    paragraph_number: int
    process_step: str
    sub_process_step: str
    data_title: str
    data_tag_es: str  # Spanish tag from extraction


# ============================================================
# No-OCR guard
# ============================================================
def assert_has_text_layer(pdf: pdfplumber.PDF, min_chars: int = 30) -> None:
    total = 0
    for p in pdf.pages[:3]:
        total += len(flat(p.extract_text() or ""))
    if total < min_chars:
        raise RuntimeError(
            "Este PDF parece NO tener capa de texto (probablemente escaneado). "
            "La extracción sin OCR no puede continuar."
        )


# ============================================================
# Core extraction (same logic)
# ============================================================
def extract_from_pdf_no_ocr(pdf_path: str, product_name_override: Optional[str] = None) -> List[DataRow]:
    rows: List[DataRow] = []
    data_no = 0

    process_step = norm(product_name_override) if product_name_override and norm(product_name_override) else "--"
    sub_process_step = ""

    table_settings = {
        "vertical_strategy": "lines",
        "horizontal_strategy": "lines",
        "intersection_tolerance": 5,
        "snap_tolerance": 3,
        "join_tolerance": 3,
        "edge_min_length": 15,
        "min_words_vertical": 1,
        "min_words_horizontal": 1,
        "text_tolerance": 3,
    }

    with pdfplumber.open(pdf_path) as pdf:
        assert_has_text_layer(pdf)

        for page_idx, page in enumerate(pdf.pages, start=1):
            text = page.extract_text() or ""
            lines = [flat(l) for l in text.split("\n") if flat(l)]

            # If UI didn't provide product_name, try to infer from document text
            if process_step == "--":
                process_step = extract_product_name(lines, fallback=process_step)

            step = extract_subprocess_step(lines)
            if step:
                sub_process_step = step

            page_split_x = compute_split_x_from_headers(page)

            # ---- tables ----
            table_objs = page.find_tables(table_settings=table_settings) or []
            for table_obj in table_objs:
                split_x = compute_split_x_from_table(table_obj) or page_split_x
                extracted = table_obj.extract() or []

                for r_i, row in enumerate(extracted, start=1):
                    items = pick_titles_with_col(row)
                    if not items:
                        continue

                    row_bbox = None
                    if hasattr(table_obj, "rows") and (r_i - 1) < len(table_obj.rows):
                        row_bbox = table_obj.rows[r_i - 1].bbox

                    for title, col_i in items:
                        if not is_meaningful_text(title):
                            continue

                        if row_bbox is not None:
                            est_x = estimate_text_x0_in_row(page, row_bbox, title)
                            data_tag_es = assign_tag_es_from_x(est_x, split_x, col_i)
                        else:
                            if split_x is None:
                                data_tag_es = TAG_TEXT_LIBRE_ES
                            else:
                                data_tag_es = TAG_LEFT_ES if col_i == 0 else TAG_RIGHT_ES

                        data_no += 1
                        rows.append(
                            DataRow(
                                data_number=data_no,
                                page_number=page_idx,
                                paragraph_number=r_i,
                                process_step=process_step,
                                sub_process_step=sub_process_step if sub_process_step else "--",
                                data_title=title,
                                data_tag_es=data_tag_es,
                            )
                        )

            # ---- fallback word rows ----
            for li, (line, row_x0) in enumerate(extract_rows_from_words(page), start=1):
                if not is_meaningful_text(line):
                    continue

                if page_split_x is None:
                    data_tag_es = TAG_TEXT_LIBRE_ES
                else:
                    data_tag_es = TAG_LEFT_ES if row_x0 < page_split_x else TAG_RIGHT_ES

                data_no += 1
                rows.append(
                    DataRow(
                        data_number=data_no,
                        page_number=page_idx,
                        paragraph_number=li,
                        process_step=process_step,
                        sub_process_step=sub_process_step if sub_process_step else "--",
                        data_title=line,
                        data_tag_es=data_tag_es,
                    )
                )

    return rows


# ============================================================
# UI-facing export (DataFrame)
# ============================================================
def build_ui_dataframe(rows: List[DataRow], file_name: str, product_name: Optional[str] = None) -> pd.DataFrame:
    # Re-number rows by sorted (page, paragraph) like your UI pipeline
    rows_sorted = sorted(rows, key=lambda r: (r.page_number, r.paragraph_number, r.data_number))

    out: List[Dict[str, Any]] = []
    for i, r in enumerate(rows_sorted, start=1):
        tag_ui = TAG_ES_TO_UI.get(r.data_tag_es, "TEXT")
        out.append({
            "Unnamed: 0": file_name if i == 1 else None,
            "Data number": i,
            "Page number": int(r.page_number),
            "Paragraph number": int(r.paragraph_number),
            "Process step": norm(product_name) if product_name and norm(product_name) else r.process_step,
            "Sub process step": r.sub_process_step if r.sub_process_step else "--",
            "Data title": r.data_title,
            "Tag": tag_data_safe(tag_ui),
        })

    df = pd.DataFrame(out, columns=COLUMNS)

    # Ensure all required columns exist
    for col in COLUMNS:
        if col not in df.columns:
            df[col] = None
    df = df[COLUMNS].copy()
    return df


# ============================================================
# Streamlit-compatible entrypoint (same as your UI structure)
# ============================================================
def extract(file_bytes: bytes, file_name: str, product_name: Optional[str] = None) -> pd.DataFrame:
    ext = os.path.splitext(file_name or "")[1].lower()
    if ext != ".pdf":
        raise ValueError("Extractor expects a PDF file.")

    with tempfile.TemporaryDirectory() as td:
        pdf_path = os.path.join(td, "input.pdf")
        with open(pdf_path, "wb") as f:
            f.write(file_bytes)

        rows = extract_from_pdf_no_ocr(pdf_path, product_name_override=product_name)
        ui_df = build_ui_dataframe(rows, file_name=file_name, product_name=product_name)
        return ui_df


# ============================================================
# Optional: local test run (not used in UI)
# ============================================================
if __name__ == "__main__":
    # Example local usage:
    # pdf = Path("/mnt/data/_1_PERINDARG WET_9490 PERINDARG 6S TO MOF 242-6 (Jun25) (2).pdf")
    # b = pdf.read_bytes()
    # df = extract(b, pdf.name, product_name="9490 PERINDARG 6S TO")
    # print(df.head(30))
    pass