import os
import re
import tempfile
from dataclasses import dataclass
from typing import Any, Dict, List, Optional, Tuple

import pdfplumber
import pandas as pd


# ============================================================
# UI OUTPUT COLUMNS (Streamlit-friendly)
# ============================================================
UI_COLUMNS = [
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
# TAGS (Spanish/French template — keep as-is)
# ============================================================
TAG_LEFT = "Consignes"
TAG_RIGHT = "Relevés paramètres"
TAG_TEXT_LIBRE = "Texte libre"

# ============================================================
# IGNORE / FILTERS (same extraction logic)
# ============================================================
IGNORE_TOKENS = {
    "", "-", "—", "↓", "→",
    "x", "✓", "☒", "☐",
    "oui", "non", "na", "n/a", "fait", "en cours", "poursuivre",
    "kg", "l", "°c", "tr/min", "min", "h",
}

IGNORE_HEADERS = {
    "consignes / champs opératoires",
    "consignes",
    "champs opératoires",
    "relevés paramètres",
    "releves parametres",
    "paramètres suivis",
    "parametres suivis",
    "action réalisée",
    "action realisee",
    "date / heure",
    "date / heure (supervision)",
    "visa",
}

SECTION_HINTS = [
    "PAGE DE GARDE",
    "MAITRISE DE LA CONTAMINATION",
    "VERIFICATION PRELIMINAIRE",
    "CONDITIONS DE DEMARRAGE",
    "PREPARATION SOLUTION",
    "TABLEAU DE RELEVE DE PARAMETRES",
    "NETTOYAGE DU MATERIEL",
]

REWRITE_RULES = [
    (re.compile(r"^Rédigé par.*", re.IGNORECASE), "Rédigé par :\n- Date et signature"),
    (re.compile(r"^Approuvé par.*", re.IGNORECASE), "Approuvé par :\n- Date et signature"),
    (re.compile(r"^Autorisé par.*", re.IGNORECASE), "Autorisé par :\n- Date et signature"),
]

_re_multispace = re.compile(r"\s+")

HEADER_NOISE_RX = re.compile(
    r"^(MB Référence|Relevé de Paramètres|Page\s+\d+\s+sur\s+\d+)\b",
    re.IGNORECASE,
)

SUB_STEP_RX = re.compile(r"^\s*\d+\.\d+\.")
MAIN_SECTION_RX = re.compile(r"^\s*(\d+)\.\s+(.*\S)\s*$")
SUBSECTION_RX = re.compile(r"^\s*\d+\.\d+\.")

UNITS_RX = re.compile(r"\b(kg|g|mg|l|ml|°c|c|min|h|tr/min|rpm|bar)\b", re.IGNORECASE)
DOTLEADER_RX = re.compile(r"\.{6,}|…{2,}")

CHAPTER_KEYWORDS = (
    "CONDITIONS DE DEMARRAGE",
    "PREPARATION SOLUTION",
    "9490-6",
    "PERINDARG",
    "SOL AC.",
    "HUMIDE",
    "PERINDARG 6S SOL",
    "PERINDARG 6S HUMIDE",
)

LABEL_VALUE_RX = re.compile(r"^(.{2,80}?)\s*:\s*(.+)$")


def norm(s: str) -> str:
    s = (s or "").replace("\u00a0", " ")
    s = s.replace("\r\n", "\n").replace("\r", "\n")
    lines = []
    for line in s.split("\n"):
        line = _re_multispace.sub(" ", line).strip()
        if line:
            lines.append(line)
    return "\n".join(lines).strip()


def flat(s: str) -> str:
    return _re_multispace.sub(" ", (s or "").replace("\u00a0", " ")).strip()


def is_noise_token(text: str) -> bool:
    t = flat(text).lower()
    if t in IGNORE_TOKENS:
        return True
    return len(t) < 2


def is_ignorable_header(text: str) -> bool:
    return flat(text).lower() in IGNORE_HEADERS


def looks_like_section_heading(text: str) -> bool:
    t = flat(text)
    if not t:
        return False
    up = t.upper()
    if any(h in up for h in SECTION_HINTS):
        return True
    if len(t) >= 10 and t == t.upper() and any(ch.isalpha() for ch in t):
        return True
    return False


def apply_rewrites(text: str) -> str:
    for rx, repl in REWRITE_RULES:
        if rx.match(text):
            return repl
    return text


def is_header_noise(line: str) -> bool:
    return bool(HEADER_NOISE_RX.match(flat(line)))


def pick_titles_with_col(row: List[Optional[str]]) -> List[Tuple[str, int]]:
    out: List[Tuple[str, int]] = []
    for col_i, cell in enumerate(row or []):
        c = norm(cell or "")
        if not c:
            continue
        if is_noise_token(c) or is_ignorable_header(c):
            continue
        c2 = apply_rewrites(c)
        if not c2 or is_noise_token(c2) or is_ignorable_header(c2):
            continue
        out.append((c2, col_i))

    seen = set()
    uniq: List[Tuple[str, int]] = []
    for text, col_i in out:
        k = text.strip().lower()
        if k not in seen:
            uniq.append((text, col_i))
            seen.add(k)
    return uniq


def extract_process_step(lines: List[str], fallback: str = "Process") -> str:
    for l in lines:
        u = flat(l).upper()
        if "PERINDARG 6S HUMIDE" in u:
            return "PERINDARG 6S HUMIDE"
        if "PERINDARG" in u and "HUMIDE" in u:
            return "PERINDARG 6S HUMIDE"
    return fallback


def extract_main_sub_process_step(lines: List[str]) -> Optional[str]:
    for l in lines:
        l = flat(l)
        if not l or is_header_noise(l):
            continue

        if SUBSECTION_RX.match(l) or SUB_STEP_RX.match(l):
            continue

        m = MAIN_SECTION_RX.match(l)
        if not m:
            continue

        n = int(m.group(1))
        title = m.group(2).strip()

        if n not in {1, 2, 3, 4, 5}:
            continue
        if DOTLEADER_RX.search(l):
            continue
        if UNITS_RX.search(l):
            continue

        up = title.upper()
        if re.search(r"\b\d+\b", title) and "9490-6" not in up:
            continue
        if not any(k in up for k in CHAPTER_KEYWORDS):
            continue

        return f"{n}. {title}"
    return None


def compute_split_x_from_table(table_obj) -> Optional[float]:
    try:
        cols = list(table_obj.columns)
        if len(cols) < 2:
            return None
        centers = []
        for col in cols:
            x0, _, x1, _ = col.bbox
            centers.append((x0 + x1) / 2.0)
        centers = sorted(centers)
        gaps = [(centers[i + 1] - centers[i], i) for i in range(len(centers) - 1)]
        if not gaps:
            return None
        max_gap, idx = max(gaps, key=lambda t: t[0])
        if max_gap < 10:
            return None
        return (centers[idx] + centers[idx + 1]) / 2.0
    except Exception:
        return None


def get_col_center_x(table_obj, col_i: int) -> Optional[float]:
    try:
        col = list(table_obj.columns)[col_i]
        x0, _, x1, _ = col.bbox
        return (x0 + x1) / 2.0
    except Exception:
        return None


def assign_data_tag_by_x(table_obj, col_i: int, fallback_col_rule: bool = True) -> str:
    split_x = compute_split_x_from_table(table_obj)
    cx = get_col_center_x(table_obj, col_i)

    if split_x is not None and cx is not None:
        return TAG_LEFT if cx < split_x else TAG_RIGHT

    if fallback_col_rule:
        return TAG_LEFT if col_i == 0 else TAG_RIGHT

    return TAG_TEXT_LIBRE


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


def extract_rows_from_words(page) -> List[str]:
    words = page.extract_words(keep_blank_chars=False, use_text_flow=True) or []
    rows = cluster_words_into_rows(words, y_tol=3.0)
    out = []
    for rw in rows:
        txt = row_text(rw)
        if not txt:
            continue
        if is_header_noise(txt):
            continue
        if is_ignorable_header(txt):
            continue
        out.append(txt)
    return out


@dataclass
class DataRow:
    data_number: int
    page_number: int
    paragraph_number: int
    process_step: str
    sub_process_step: str
    data_title: str
    data_tag: str


# ============================================================
# EXTRACTION (same logic; only paragraph renumbering + UI override)
# ============================================================
def extract_from_pdf(
    pdf_path: str,
    product_name_override: Optional[str] = None,
) -> List[DataRow]:
    rows: List[DataRow] = []
    data_no = 0

    # keep same behavior: UI override if present, otherwise fallback detection
    process_step = norm(product_name_override) if product_name_override and norm(product_name_override) else "Process"
    sub_process_step = ""

    seen_titles_per_page: Dict[Tuple[int, str], int] = {}

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
        for page_idx, page in enumerate(pdf.pages, start=1):
            # ✅ NEW: continuous paragraph numbering per page (no gaps)
            para_counter = 0

            text = page.extract_text() or ""
            lines = [flat(l) for l in text.split("\n") if flat(l)]

            # Only auto-extract when UI didn't provide it (same logic)
            if not (product_name_override and norm(product_name_override)):
                process_step = extract_process_step(lines, fallback=process_step)

            main_step = extract_main_sub_process_step(lines)
            if main_step:
                sub_process_step = main_step

            table_objs = page.find_tables(table_settings=table_settings) or []
            for table_obj in table_objs:
                extracted = table_obj.extract() or []
                for _r_i, row in enumerate(extracted, start=1):
                    items = pick_titles_with_col(row)
                    if not items:
                        continue

                    for title, col_i in items:
                        key = (page_idx, title.strip().lower())
                        seen_titles_per_page[key] = seen_titles_per_page.get(key, 0) + 1

                        data_tag = assign_data_tag_by_x(table_obj, col_i)

                        data_no += 1
                        para_counter += 1  # ✅ only increments when we actually keep a row
                        rows.append(
                            DataRow(
                                data_number=data_no,
                                page_number=page_idx,
                                paragraph_number=para_counter,
                                process_step=process_step,
                                sub_process_step=sub_process_step,
                                data_title=title,
                                data_tag=data_tag,
                            )
                        )

            word_rows = extract_rows_from_words(page)
            for _li, line in enumerate(word_rows, start=1):
                if not line or is_header_noise(line):
                    continue
                if is_noise_token(line) or is_ignorable_header(line):
                    continue

                is_label_value = bool(LABEL_VALUE_RX.match(line))
                is_heading = looks_like_section_heading(line)
                is_meaningful = (len(line) >= 12 and sum(ch.isalpha() for ch in line) >= 5)

                if not (is_label_value or is_heading or is_meaningful):
                    continue

                key = (page_idx, line.strip().lower())
                if key in seen_titles_per_page:
                    continue

                if line.lower() in IGNORE_HEADERS:
                    continue

                data_no += 1
                para_counter += 1  # ✅ only increments when we actually keep a row
                rows.append(
                    DataRow(
                        data_number=data_no,
                        page_number=page_idx,
                        paragraph_number=para_counter,
                        process_step=process_step,
                        sub_process_step=sub_process_step,
                        data_title=line,
                        data_tag=TAG_TEXT_LIBRE,
                    )
                )

    return rows


def extract_to_dataframe(pdf_path: str, product_name: Optional[str] = None) -> pd.DataFrame:
    rows = extract_from_pdf(pdf_path, product_name_override=product_name)

    data = []
    for r in rows:
        data.append({
            "Unnamed: 0": None,  # filled later (first row only)
            "Data number": r.data_number,
            "Page number": r.page_number,
            "Paragraph number": r.paragraph_number,
            "Process step": r.process_step,
            "Sub process step": r.sub_process_step if r.sub_process_step else "--",
            "Data title": r.data_title,
            "Tag": r.data_tag,
        })

    return pd.DataFrame(data, columns=UI_COLUMNS)


def extract(file_bytes: bytes, file_name: str, product_name: Optional[str] = None) -> pd.DataFrame:
    ext = os.path.splitext(file_name or "")[1].lower()
    if ext != ".pdf":
        raise ValueError("Extractor expects a PDF file.")

    with tempfile.TemporaryDirectory() as td:
        pdf_path = os.path.join(td, "input.pdf")
        with open(pdf_path, "wb") as f:
            f.write(file_bytes)

        ui_df = extract_to_dataframe(pdf_path, product_name=product_name)

        # Unnamed: 0 should be file name only on first row
        ui_df["Unnamed: 0"] = None
        if len(ui_df) > 0:
            ui_df.iloc[0, ui_df.columns.get_loc("Unnamed: 0")] = file_name

        # ✅ Force Process step from UI (always), no dependency on PDF detection
        ps = norm(product_name) if product_name and str(product_name).strip() else "--"
        ui_df["Process step"] = ps

        return ui_df[UI_COLUMNS].copy()