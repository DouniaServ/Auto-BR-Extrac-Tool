# -*- coding: utf-8 -*-

import os
import re
import tempfile
from typing import Dict, List, Optional, Tuple

import fitz  # PyMuPDF
import pandas as pd


# =============================================================================
# Utilities
# =============================================================================

def norm(s: str) -> str:
    return re.sub(r"\s+", " ", (s or "").strip())

def strip_accents_lower(s: str) -> str:
    s = (s or "").lower().strip()
    return (s.replace("é", "e").replace("è", "e").replace("ê", "e")
            .replace("à", "a").replace("ù", "u").replace("î", "i")
            .replace("’", "'").replace("‘", "'"))

def normalize_for_repetition(s: str) -> str:
    t = strip_accents_lower(norm(s))
    t = re.sub(r"\bpage\b\s*\d+\s*/\s*\d+", "page#/total#", t)
    t = re.sub(r"\bpage\b\s*\d+", "page#", t)
    t = re.sub(r"\d{2}/\d{2}/\d{4}", "date#", t)
    t = re.sub(r"\d{4}-\d{2}-\d{2}", "date#", t)
    t = re.sub(r"\b\d{1,2}:\d{2}\b", "time#", t)
    return t


# =============================================================================
# Skip noise: copie / travail / copie de travail + skip pure "de"
# =============================================================================

SKIP_INLINE_RX = re.compile(r"(?i)\b(copie\s+de\s+travail|copie|travail)\b")
PURE_NOISE_RX  = re.compile(r"(?i)^\s*(copie(\s+de\s+travail)?|travail)\s*$")
PURE_DE_RX     = re.compile(r"(?i)^\s*de\s*$")  # skip lines that are only "de"

def remove_skip_noise_inline(text: str) -> str:
    if not text:
        return ""
    cleaned = SKIP_INLINE_RX.sub("", text)
    return norm(cleaned)

def is_pure_noise_line(s: str) -> bool:
    t = norm(s or "")
    if not t:
        return True
    if PURE_DE_RX.match(t):
        return True
    low = strip_accents_lower(t)
    if "conformite" in low and "edition" in low:
        return False
    return bool(PURE_NOISE_RX.match(t))

def clean_row_noise(r: Dict[str, str]) -> Dict[str, str]:
    r["Data Title"] = remove_skip_noise_inline(r.get("Data Title", ""))
    r["Théorique"]  = remove_skip_noise_inline(r.get("Théorique", ""))
    r["Réel"]       = remove_skip_noise_inline(r.get("Réel", ""))
    r["Visa"]       = remove_skip_noise_inline(r.get("Visa", ""))
    r["N° observ."] = remove_skip_noise_inline(r.get("N° observ.", ""))
    return r

def should_skip_text(text: str) -> bool:
    return is_pure_noise_line(text)


# =============================================================================
# CONFORMITE DE L’EDITION
# =============================================================================

CONFORMITE_EDITION_LABEL = "CONFORMITE DE L’EDITION"

def split_conformite_on_date_visa(text: str) -> Optional[List[str]]:
    s = norm(text)
    low = strip_accents_lower(s)
    if not ("conformite" in low and "edition" in low):
        return None

    has_date = "date" in low
    has_visa = "visa" in low
    has_on = (
        ("o / n" in low) or ("o/n" in low)
        or bool(re.search(r"\bo\s*/?\s*n\b", low))
        or bool(re.search(r"\bo\s+n\b", low))
    )

    if has_date and has_visa and has_on:
        return [CONFORMITE_EDITION_LABEL, "O / N", "Date", "Visa"]
    if has_date and has_visa:
        return [CONFORMITE_EDITION_LABEL, "Date", "Visa"]
    return [CONFORMITE_EDITION_LABEL]

def normalize_conformite_header(title: str) -> str:
    s = norm(title)
    low = strip_accents_lower(s)
    if "conformite" in low and "edition" in low:
        return CONFORMITE_EDITION_LABEL
    return s


# =============================================================================
# PROCESS STEP / SUB PROCESS STEP detection
# =============================================================================

MAIN_STAGE_KEYWORDS = [
    ("granulation", "GRANULATION"),
    ("compression", "COMPRESSION"),
    ("pelliculage", "PELLICULAGE"),
    ("enrobage", "ENROBAGE"),
]

MAIN_STAGE_EXCLUDE = [
    "compte rendu", "fabrication", "semi-fini",
    "operateurs", "opérateurs",
    "conformite", "conformité", "edition", "édition",
    "drf", "action", "n°", "no", "observ", "visa", "theorique", "reel"
]

def normalize_main_stage(text: str) -> Optional[str]:
    low = strip_accents_lower(text)
    for kw, label in MAIN_STAGE_KEYWORDS:
        if kw in low:
            if label == "GRANULATION" and "humide" in low:
                return "GRANULATION HUMIDE"
            return label
    return None

NUMBERED_HDR_RX = re.compile(r"^\s*\d+\s*[–-]\s*\S+", re.UNICODE)

def is_numbered_header(text: str) -> bool:
    return bool(NUMBERED_HDR_RX.match(norm(text)))

def normalize_numbered_header(text: str) -> str:
    return norm(text).upper()

TRAILING_DE_RX = re.compile(r"(?i)\s+\bde\b\s*$")
def strip_trailing_de_on_headers(text: str) -> str:
    t = norm(text)
    if is_numbered_header(t):
        t = TRAILING_DE_RX.sub("", t)
    return norm(t)

def update_subprocess_step(current: str, text: str) -> str:
    s = norm(text)
    if not s or should_skip_text(s):
        return current

    low = strip_accents_lower(s)
    if "operateurs" in low and "participants" in low:
        return "OPERATEURS PARTICIPANTS"

    if "conformite" in low and "edition" in low:
        return CONFORMITE_EDITION_LABEL

    if is_numbered_header(s):
        return normalize_numbered_header(strip_trailing_de_on_headers(s))

    return current


# =============================================================================
# Header/footer removal (two-pass)
# =============================================================================

def get_blocks(page: fitz.Page) -> List[Tuple[float, float, float, float, str]]:
    blocks = page.get_text("blocks")
    out = []
    for b in blocks:
        x0, y0, x1, y1, text = b[0], b[1], b[2], b[3], b[4]
        text = norm(text)
        if text:
            out.append((float(x0), float(y0), float(x1), float(y1), text))
    out.sort(key=lambda t: (t[1], t[0]))
    return out

def detect_main_stage(page: fitz.Page, top_band: float = 260) -> Optional[str]:
    candidates: List[Tuple[float, float, str]] = []
    for x0, y0, x1, y1, txt in get_blocks(page):
        if y0 > top_band:
            break

        t = norm(txt)
        if not t:
            continue

        low = strip_accents_lower(t)
        label = normalize_main_stage(t)
        if not label:
            continue
        if any(bad in low for bad in MAIN_STAGE_EXCLUDE):
            continue

        candidates.append((float(y0), float(x0), label))

    if not candidates:
        return None
    candidates.sort(key=lambda z: (z[0], z[1]))
    return candidates[0][2]

def learn_repeating_header_footer(
    doc: fitz.Document,
    top_band: float = 120,
    bottom_band: float = 120,
    min_len: int = 12,
    min_ratio: float = 0.30
) -> set:
    counts: Dict[str, int] = {}
    n_pages = doc.page_count
    for i in range(n_pages):
        page = doc.load_page(i)
        H = float(page.rect.height)
        for x0, y0, x1, y1, txt in get_blocks(page):
            if y0 <= top_band or y1 >= (H - bottom_band):
                key = normalize_for_repetition(txt)
                if len(key) >= min_len:
                    counts[key] = counts.get(key, 0) + 1
    threshold = max(2, int(n_pages * min_ratio))
    return {k for k, c in counts.items() if c >= threshold}


# =============================================================================
# Detect page title subprocess (Page 1: COMPTE RENDU...)
# =============================================================================

SEMI_FINI_RX = re.compile(r"(?i)\bcompte\s+rendu\s+de\s+fabrication\s+du\s+semi[-\s]?fini\b")

def detect_cover_subprocess(page: fitz.Page, top_band: float = 240) -> Optional[str]:
    for x0, y0, x1, y1, txt in get_blocks(page):
        if y0 > top_band:
            break
        if SEMI_FINI_RX.search(txt):
            return "COMPTE RENDU DE FABRICATION DU SEMI-FINI"
    return None


# =============================================================================
# Word extraction (geometry)
# =============================================================================

def get_words(page: fitz.Page, header_margin: float = 60, footer_margin: float = 60) -> List[dict]:
    H = float(page.rect.height)
    raw = page.get_text("words")

    out = []
    for x0, y0, x1, y1, text, *_ in raw:
        if y0 < header_margin:
            continue
        if y1 > (H - footer_margin):
            continue
        text = norm(text)
        if not text:
            continue
        out.append({
            "text": text,
            "x0": float(x0), "x1": float(x1),
            "top": float(y0), "bottom": float(y1),
        })
    out.sort(key=lambda w: (w["top"], w["x0"]))
    return out


# =============================================================================
# Table detection + boundaries
# =============================================================================

def header_key(text: str) -> str:
    t = strip_accents_lower(norm(text))
    if t in ("theorique", "théorique"):
        return "theorique"
    if t in ("reel", "réel"):
        return "reel"
    if t == "visa":
        return "visa"
    if t in ("n°", "nº", "n", "no", "observ", "observ.", "n°observ", "n°observ."):
        return "observ"
    return ""

def find_header_and_columns(
    words: List[dict],
    page_height: float,
    y_band: float = 3.0,
    max_header_y_ratio: float = 0.45
) -> Optional[Dict[str, float]]:
    bands: Dict[int, List[Tuple[str, float, float, float]]] = {}
    for w in words:
        k = header_key(w["text"])
        if not k:
            continue
        band = round(w["top"] / y_band)
        xc = (w["x0"] + w["x1"]) / 2.0
        bands.setdefault(band, []).append((k, xc, w["top"], w["bottom"]))

    if not bands:
        return None

    best_band = None
    best_keys = set()
    for b, items in bands.items():
        keys = set(k for k, *_ in items)
        band_y = min(top for _, _, top, _ in items)
        if band_y > page_height * max_header_y_ratio:
            continue
        if len(keys) > len(best_keys):
            best_keys = keys
            best_band = b

    required = {"theorique", "reel", "visa"}
    if best_band is None or not required.issubset(best_keys):
        return None

    cols: Dict[str, float] = {}
    bottoms = []
    for k, xc, top, bottom in bands[best_band]:
        cols[k] = min(cols.get(k, 1e9), xc)
        bottoms.append(bottom)

    if "observ" not in cols:
        cols["observ"] = max(xc for _, xc, *_ in bands[best_band])

    cols["_table_top"] = max(bottoms) + 3.0
    return cols

def make_x_boundaries(cols: Dict[str, float], page_width: float) -> Dict[str, Tuple[float, float]]:
    xs = sorted([cols["theorique"], cols["reel"], cols["visa"], cols["observ"]])
    x_th, x_re, x_vi, x_ob = xs
    return {
        "title":     (0,    x_th),
        "theorique": (x_th, x_re),
        "reel":      (x_re, x_vi),
        "visa":      (x_vi, x_ob),
        "observ":    (x_ob, page_width),
    }


# =============================================================================
# Line clustering + column text
# =============================================================================

def cluster_words_by_line(words: List[dict], y_tol: float = 3.0) -> List[List[dict]]:
    words = sorted(words, key=lambda w: (w["top"], w["x0"]))
    lines: List[List[dict]] = []
    cur: List[dict] = []
    cur_y: Optional[float] = None

    for w in words:
        if cur_y is None:
            cur = [w]
            cur_y = w["top"]
            continue
        if abs(w["top"] - cur_y) <= y_tol:
            cur.append(w)
            cur_y = (cur_y + w["top"]) / 2.0
        else:
            lines.append(cur)
            cur = [w]
            cur_y = w["top"]

    if cur:
        lines.append(cur)
    return lines

def words_in_range_with_x(line_words: List[dict], x_min: float, x_max: float) -> Tuple[str, Optional[float]]:
    ws = [w for w in line_words if x_min <= ((w["x0"] + w["x1"]) / 2.0) < x_max]
    ws.sort(key=lambda w: w["x0"])
    txt = norm(" ".join(w["text"] for w in ws))
    min_x0 = min([w["x0"] for w in ws], default=None)
    return txt, min_x0


# =============================================================================
# Post-processing merges
# =============================================================================

END_DANGLING_RX = re.compile(
    r"""(?ix)
    (?:\b(et|ou|de|du|des|la|le|les|un|une|dans|au|aux|pour|avec|sans|par|sur|en|a)\b\s*$)
    |(?:[/:,-]\s*$)
    """
)
TERMINAL_PUNCT_RX = re.compile(r"[.!?;:]$")
STARTS_LOWER_RX = re.compile(r"^[a-zà-ÿ]")
LABEL_END_RX = re.compile(r"(?:→|->|:)\s*$")

VALUE_LIKE_RX = re.compile(
    r"""(?ix)
    ^\s*(
        on|off|oui|non|ok|ko|c\s*/\s*nc
        |[<>≤≥]?\s*\d+(?:[.,]\d+)?\s*(?:±\s*\d+(?:[.,]\d+)?)?\s*(?:\([^)]*\))?
        |\d+(?:[.,]\d+)?\s*(?:kg|g|mg|mbar|bar|°c|c|%|min|h|tr/min|rpm)\b
    )\s*$
    """
)

BULLET_RX = re.compile(r"^\s*[-•‣▪●・\u2022\u25AA\u25CF\uF0B7]\s+(.*\S)\s*$")
LEADING_BULLETS_RX = re.compile(r"^\s*[-•‣▪●・\u2022\u25AA\u25CF\uF0B7]+\s*")

def normalize_label_line(s: str) -> str:
    s = norm(s)
    return LEADING_BULLETS_RX.sub("", s)

def is_resources_label(s: str) -> bool:
    s = normalize_label_line(s)
    low = strip_accents_lower(s)
    return low.startswith("ressources utilisees") or low.startswith("ressource utilisee")

def _starts_continuation(s: str) -> bool:
    s = norm(s)
    if not s:
        return False
    return bool(STARTS_LOWER_RX.match(s)) or s.lower().startswith(("et ", "ou ", "de ", "du ", "des "))

def _is_value_like(s: str) -> bool:
    s = norm(s)
    if not s or len(s) > 40:
        return False
    return bool(VALUE_LIKE_RX.match(s))

def _is_label_like(s: str) -> bool:
    s = norm(s)
    if not s:
        return False
    if any(ch.isdigit() for ch in s):
        return False
    return len(s) <= 80

def merge_wrapped_titles(rows: List[Dict[str, str]]) -> List[Dict[str, str]]:
    out: List[Dict[str, str]] = []
    MAX_JOINED_LEN_WRAP = 220
    MAX_JOINED_LEN_LABELVAL = 160

    i = 0
    while i < len(rows):
        r = rows[i]

        # label + value merge
        if i + 1 < len(rows):
            n = rows[i + 1]

            r_title_only = bool(r.get("Data Title")) and not (r.get("Théorique") or r.get("Réel") or r.get("Visa") or r.get("N° observ."))
            n_title_only = bool(n.get("Data Title")) and not (n.get("Théorique") or n.get("Réel") or n.get("Visa") or n.get("N° observ."))

            if r_title_only and n_title_only:
                t1 = norm(r.get("Data Title", ""))
                t2 = norm(n.get("Data Title", ""))

                px = r.get("_title_x0")
                cx = n.get("_title_x0")
                same_indent = (px is not None and cx is not None and cx >= px - 2)

                if same_indent and (LABEL_END_RX.search(t1) or _is_label_like(t1)) and _is_value_like(t2):
                    label = LABEL_END_RX.sub("", t1).strip()
                    merged = norm(f"{label} ({t2})")
                    if len(merged) <= MAX_JOINED_LEN_LABELVAL:
                        r["Data Title"] = merged
                        i += 1

        if not out:
            out.append(r)
            i += 1
            continue

        # wrapped sentence merge
        prev = out[-1]
        title_only = bool(r.get("Data Title")) and not (r.get("Théorique") or r.get("Réel") or r.get("Visa") or r.get("N° observ."))
        prev_title_only = bool(prev.get("Data Title")) and not (prev.get("Théorique") or prev.get("Réel") or prev.get("Visa") or prev.get("N° observ."))

        if title_only and prev_title_only:
            px = prev.get("_title_x0")
            cx = r.get("_title_x0")
            prev_txt = norm(prev.get("Data Title", ""))
            cur_txt = norm(r.get("Data Title", ""))

            same_indent = (px is not None and cx is not None and cx >= px - 2)
            looks_like_wrap = bool(END_DANGLING_RX.search(prev_txt)) and _starts_continuation(cur_txt)
            prev_not_ended = not bool(TERMINAL_PUNCT_RX.search(prev_txt))

            if same_indent and (looks_like_wrap or (prev_not_ended and _starts_continuation(cur_txt))):
                joined = norm(prev_txt + " " + cur_txt)
                if len(joined) <= MAX_JOINED_LEN_WRAP:
                    prev["Data Title"] = joined
                    i += 1
                    continue

        out.append(r)
        i += 1

    return out

def merge_bullets_under_resources(rows: List[Dict[str, str]]) -> List[Dict[str, str]]:
    """
    Merge:
      ' Ressources utilisées :' (or 'Ressources utilisées :')
      followed by bullet rows ('- xxx', '• xxx', ' xxx', ...)
    into ONE row.
    """
    out: List[Dict[str, str]] = []
    i = 0

    while i < len(rows):
        r = rows[i]
        raw_title = norm(r.get("Data Title", ""))
        label_title = normalize_label_line(raw_title)

        r_title_only = bool(raw_title) and not (r.get("Théorique") or r.get("Réel") or r.get("Visa") or r.get("N° observ."))
        if r_title_only and is_resources_label(raw_title) and label_title.endswith(":"):
            bullets: List[str] = []
            j = i + 1

            while j < len(rows):
                n = rows[j]
                n_title = norm(n.get("Data Title", ""))
                n_title_only = bool(n_title) and not (n.get("Théorique") or n.get("Réel") or n.get("Visa") or n.get("N° observ."))

                if not n_title_only:
                    break
                if is_numbered_header(n_title):
                    break

                m = BULLET_RX.match(n_title)
                if not m:
                    break

                bullets.append(norm(m.group(1)))
                j += 1

            if bullets:
                r["Data Title"] = label_title + " " + "; ".join(bullets)
                out.append(r)
                i = j
                continue

        out.append(r)
        i += 1

    return out

NUM_ONLY_RX = re.compile(r"^\s*\d+(?:[.,]\d+)?\s*$")
STARTS_WITH_UNITISH_RX = re.compile(r"(?i)^\s*(%|kg|g|mg|mbar|bar|°\s*c|°c|c|min|h|tr/min|rpm)\b")

def merge_orphan_value_rows(rows: List[Dict[str, str]]) -> List[Dict[str, str]]:
    """
    Merge into previous row when:
      A) Data Title is empty but value cols exist
      B) OR Data Title is numeric-only (ex: '99,0') and Théorique starts with a unit/percent (ex: '% 101,0%')
         => produce '99,0 % 101,0%' in previous row Théorique.
    """
    out: List[Dict[str, str]] = []

    for r in rows:
        title = norm(r.get("Data Title", ""))
        th = norm(r.get("Théorique", ""))
        re_ = norm(r.get("Réel", ""))
        vi = norm(r.get("Visa", ""))
        ob = norm(r.get("N° observ.", ""))

        has_any_value = bool(th or re_ or vi or ob)

        if out:
            prev = out[-1]

            # (B) numeric-only title + theorique fragment like "% 101,0%" -> merge into previous row theorique
            if title and NUM_ONLY_RX.match(title) and th and STARTS_WITH_UNITISH_RX.match(th) and not (re_ or vi or ob):
                merged_th = norm(f"{title} {th}")
                if prev.get("Théorique"):
                    prev["Théorique"] = norm(prev.get("Théorique", "") + " " + merged_th)
                else:
                    prev["Théorique"] = merged_th
                continue

            # (A) empty title but has values -> merge into previous row
            if (not title) and has_any_value:
                if th:
                    prev["Théorique"] = norm((prev.get("Théorique", "") + " " + th).strip())
                if re_:
                    prev["Réel"] = norm((prev.get("Réel", "") + " " + re_).strip())
                if vi:
                    prev["Visa"] = norm((prev.get("Visa", "") + " " + vi).strip())
                if ob:
                    prev["N° observ."] = norm((prev.get("N° observ.", "") + " " + ob).strip())
                continue

        out.append(r)

    return out


# =============================================================================
# Extract table rows + non-table paragraphs
# =============================================================================

def extract_form_table_rows(page: fitz.Page, words: List[dict]) -> Optional[List[Dict[str, str]]]:
    H = float(page.rect.height)
    cols = find_header_and_columns(words, page_height=H)
    if not cols:
        return None

    bounds = make_x_boundaries(cols, float(page.rect.width))
    below = [w for w in words if w["top"] >= cols["_table_top"]]
    lines = cluster_words_by_line(below, y_tol=3.0)

    rows: List[Dict[str, str]] = []
    for ln in lines:
        title, title_x0 = words_in_range_with_x(ln, *bounds["title"])
        th, _ = words_in_range_with_x(ln, *bounds["theorique"])
        re_, _ = words_in_range_with_x(ln, *bounds["reel"])
        vi, _ = words_in_range_with_x(ln, *bounds["visa"])
        ob, _ = words_in_range_with_x(ln, *bounds["observ"])

        row = {
            "Data Title": strip_trailing_de_on_headers(title),
            "Théorique": th,
            "Réel": re_,
            "Visa": vi,
            "N° observ.": ob,
            "_title_x0": title_x0,
        }

        row = clean_row_noise(row)

        if is_pure_noise_line(row["Data Title"]) and not any([row["Théorique"], row["Réel"], row["Visa"], row["N° observ."]]):
            continue

        if PURE_DE_RX.match(norm(row["Data Title"])):
            continue

        rows.append(row)

    meaningful_titles = sum(1 for r in rows if r["Data Title"])
    filled = sum(1 for r in rows if (r["Théorique"] or r["Réel"] or r["Visa"] or r["N° observ."]))
    if meaningful_titles < 3 and filled < 3:
        return None

    rows = merge_wrapped_titles(rows)
    rows = merge_bullets_under_resources(rows)
    rows = merge_orphan_value_rows(rows)
    return rows

def extract_non_table_paragraphs(
    page: fitz.Page,
    repeating_keys: set,
    header_margin: float = 80,
    footer_margin: float = 80
) -> List[str]:
    paras = []
    H = float(page.rect.height)
    for x0, y0, x1, y1, txt in get_blocks(page):
        if y0 < header_margin:
            continue
        if y1 > (H - footer_margin):
            continue
        key = normalize_for_repetition(txt)
        if key in repeating_keys:
            continue
        if is_pure_noise_line(txt):
            continue

        txt = remove_skip_noise_inline(txt)
        if txt and not PURE_DE_RX.match(norm(txt)):
            paras.append(txt)
    return paras


# =============================================================================
# Data tag + Formula detection
# =============================================================================

TABLE_HEADER_HINTS = {
    "nom", "prenom", "nom prenom",
    "date", "visa", "signature", "matricule",
    "fonction", "operateur", "operateurs", "opérateurs",
    "commentaire", "commentaires",
    "change control", "change control n°", "change control n° :",
    "derogation", "dérogation", "derogation n°", "dérogation n° :",
    "conformite de l'edition", "conformité de l'édition",
    "conformite de l'etape", "conformité de l'étape",
    "drf/drf action", "drf / drf action",
}

def is_table_header_title(title: str) -> bool:
    t = strip_accents_lower(norm(title))
    if not t:
        return False
    if t in TABLE_HEADER_HINTS:
        return True
    if len(t) <= 18 and re.fullmatch(r"[a-zà-ÿ' /-]+", t):
        if any(k in t for k in ("nom", "prenom", "date", "visa", "signature", "matricule")):
            return True
    if t == "visa" or t == "date":
        return True
    return False

FORMULA_RX = re.compile(
    r"""(?ix)
    (?:[×x*]\s*\d+)
    |(?:/\s*\d+)
    |(?:\b\d+\s*[%]\b)
    """
)

def is_resources_title(title: str) -> bool:
    t = LEADING_BULLETS_RX.sub("", norm(title))
    low = strip_accents_lower(t)
    return low.startswith("ressources utilisees") or low.startswith("ressource utilisee")

def compute_data_tag(data_title: str, visa: str = "", theorique: str = "", subprocess: str = "") -> str:
    title = norm(data_title or "")
    low_title = strip_accents_lower(title)
    low_sub = strip_accents_lower(norm(subprocess))

    if is_resources_title(title):
        return "ressources"

    if "conformite" in low_sub and "edition" in low_sub:
        return "visa"

    if FORMULA_RX.search(title):
        return "Formula"

    if "etiquette" in low_title:
        return "etiquette"
    if is_table_header_title(title):
        return "table header"
    if norm(visa):
        return "visa"
    if norm(theorique):
        return "théorique"
    return "instruction"


# =============================================================================
# Théorique detection: avoid false positives like "0053"
# =============================================================================

THEO_STRONG_RX = re.compile(r"(?i)(?:±|≤|≥|<|>|%|\b(?:kg|g|mg|mbar|bar|°c|c|min|h|tr/min|rpm)\b)")
THEO_PAREN_UNIT_RX = re.compile(r"(?i)\([^)]*(?:kg|g|mg|mbar|bar|°c|c|min|h|tr/min|rpm|%)\)")
PURE_ID_RX = re.compile(r"^\s*\d{1,5}\s*$")

def is_true_theorique_value(v: str) -> bool:
    v = norm(v)
    if not v:
        return False
    if PURE_ID_RX.match(v):
        return False
    if THEO_STRONG_RX.search(v):
        return True
    if THEO_PAREN_UNIT_RX.search(v):
        return True
    return False


# =============================================================================
# FIX: If theorique is split between title and theorique column
# =============================================================================

TRAIL_VALUE_RX = re.compile(
    r"""(?ix)
    ^(?P<label>.*?)
    (?P<val>
        [<>≤≥]?\s*\d+(?:[.,]\d+)?\s*
        (?:°\s*c|°c|c|%|kg|g|mg|mbar|bar|min|h|tr/min|rpm)
    )\s*$
    """
)

THEO_CONT_RX = re.compile(r"(?ix)^\s*(?:±|[<>≤≥])")

def fix_split_theorique(title: str, theorique: str) -> Tuple[str, str]:
    t = norm(title)
    th = norm(theorique)
    if not t or not th:
        return title, theorique

    m = TRAIL_VALUE_RX.match(t)
    if not m:
        return title, theorique

    if not THEO_CONT_RX.match(th):
        return title, theorique

    label = norm(m.group("label"))
    val = norm(m.group("val"))
    merged = norm(f"{val} {th}")

    if is_true_theorique_value(merged):
        return label, merged

    return title, theorique


# =============================================================================
# Keep value columns (support human) + apply split fix
# =============================================================================

KEEP_VISA_VALUE = True

def enrich_keep_values(r: Dict[str, str], theo_col: str = "Théorique value", visa_col: str = "Visa value") -> Dict[str, str]:
    title = norm(r.get("Data Title", ""))
    th = norm(r.get("Théorique", ""))
    vi = norm(r.get("Visa", ""))
    subp = norm(r.get("Sub process step", ""))

    if title and th:
        new_title, new_th = fix_split_theorique(title, th)
        r["Data Title"] = new_title
        r["Théorique"] = new_th
        title = new_title
        th = new_th

    if theo_col not in r:
        r[theo_col] = ""
    if th:
        r[theo_col] = th

    if KEEP_VISA_VALUE:
        if visa_col not in r:
            r[visa_col] = ""
        if vi:
            r[visa_col] = vi

    low_sub = strip_accents_lower(subp)
    if "conformite" in low_sub and "edition" in low_sub:
        r["Data tag"] = "visa"
    else:
        if vi:
            r["Data tag"] = "visa"
        elif th and is_true_theorique_value(th):
            r["Data tag"] = "théorique"

    return r


# =============================================================================
# Extraction engine (UNCHANGED LOGIC) + UI-friendly wrappers
# =============================================================================

def pdf_to_dataframe(pdf_path: str, compact_output: bool = True) -> pd.DataFrame:
    """
    Same extraction logic as pdf_to_excel, but returns the DataFrame (for UI usage).
    Ensures raw columns exist for downstream validation.
    """
    pdf_name = os.path.basename(pdf_path)

    rows_out: List[Dict[str, str]] = []
    data_no = 0

    current_process = ""
    current_subprocess = ""

    doc = fitz.open(pdf_path)
    try:
        repeating = learn_repeating_header_footer(doc)

        for pi in range(doc.page_count):
            page = doc.load_page(pi)

            cover = detect_cover_subprocess(page)
            if cover and pi == 0:
                current_subprocess = cover

            stage = detect_main_stage(page)
            if stage:
                current_process = stage

            # Page 1 => 0, Page 2+ => 1 (constant per page)
            para_no = 0 if pi == 0 else 1

            words = get_words(page, header_margin=60, footer_margin=60)
            table_rows = extract_form_table_rows(page, words)

            if table_rows:
                for r in table_rows:
                    r = clean_row_noise(r)

                    conf_parts = split_conformite_on_date_visa(r["Data Title"])
                    if conf_parts:
                        sp_out = current_subprocess if norm(current_subprocess) else CONFORMITE_EDITION_LABEL
                        for part in conf_parts:
                            data_no += 1
                            rows_out.append({
                                "PDF name": pdf_name,
                                "Data number": data_no,
                                "Page number": pi + 1,
                                "Paragraph number": para_no,
                                "Process step": current_process if norm(current_process) else "--",
                                "Sub process step": sp_out,
                                "Data Title": part,
                                "Data tag": "visa",
                                "Théorique": "",
                                "Réel": "",
                                "Visa": "",
                                "N° observ.": "",
                            })
                        continue

                    title = strip_trailing_de_on_headers(normalize_conformite_header(r["Data Title"]))

                    if is_numbered_header(title):
                        current_subprocess = normalize_numbered_header(title)

                    sp_out = current_subprocess if norm(current_subprocess) else "--"

                    data_no += 1
                    rows_out.append({
                        "PDF name": pdf_name,
                        "Data number": data_no,
                        "Page number": pi + 1,
                        "Paragraph number": para_no,
                        "Process step": current_process if norm(current_process) else "--",
                        "Sub process step": sp_out,
                        "Data Title": title,
                        "Data tag": compute_data_tag(title, visa=r["Visa"], theorique=r["Théorique"], subprocess=sp_out),
                        "Théorique": r["Théorique"],
                        "Réel": r["Réel"],
                        "Visa": r["Visa"],
                        "N° observ.": r["N° observ."],
                    })
                continue

            paras = extract_non_table_paragraphs(page, repeating_keys=repeating)
            for txt in paras:
                txt = remove_skip_noise_inline(txt)
                if not txt or PURE_DE_RX.match(norm(txt)):
                    continue

                conf_parts = split_conformite_on_date_visa(txt)
                if conf_parts:
                    sp_out = current_subprocess if norm(current_subprocess) else CONFORMITE_EDITION_LABEL
                    for part in conf_parts:
                        data_no += 1
                        rows_out.append({
                            "PDF name": pdf_name,
                            "Data number": data_no,
                            "Page number": pi + 1,
                            "Paragraph number": para_no,
                            "Process step": current_process if norm(current_process) else "--",
                            "Sub process step": sp_out,
                            "Data Title": part,
                            "Data tag": "visa",
                            "Théorique": "",
                            "Réel": "",
                            "Visa": "",
                            "N° observ.": "",
                        })
                    continue

                txt = strip_trailing_de_on_headers(normalize_conformite_header(txt))

                if is_numbered_header(txt):
                    current_subprocess = normalize_numbered_header(txt)

                sp_out = current_subprocess if norm(current_subprocess) else "--"

                data_no += 1
                rows_out.append({
                    "PDF name": pdf_name,
                    "Data number": data_no,
                    "Page number": pi + 1,
                    "Paragraph number": para_no,
                    "Process step": current_process if norm(current_process) else "--",
                    "Sub process step": sp_out,
                    "Data Title": txt,
                    "Data tag": compute_data_tag(txt, subprocess=sp_out),
                    "Théorique": "",
                    "Réel": "",
                    "Visa": "",
                    "N° observ.": "",
                })

    finally:
        doc.close()

    df = pd.DataFrame(rows_out)

    recs = df.to_dict("records")
    recs = [enrich_keep_values(r, theo_col="Théorique value", visa_col="Visa value") for r in recs]
    df = pd.DataFrame(recs)

    # Ensure App.py-required raw columns exist (even if you later subset)
    required_raw = ["Théorique", "Réel", "Visa", "N° observ."]
    for c in required_raw:
        if c not in df.columns:
            df[c] = ""

    # Optionally populate raw from value columns when raw empty
    if "Théorique value" in df.columns:
        mask = df["Théorique"].fillna("").eq("")
        df.loc[mask, "Théorique"] = df.loc[mask, "Théorique value"].fillna("")
    if "Visa value" in df.columns:
        mask = df["Visa"].fillna("").eq("")
        df.loc[mask, "Visa"] = df.loc[mask, "Visa value"].fillna("")

    if compact_output:
        keep_cols = [
            "PDF name",
            "Data number", "Page number", "Paragraph number",
            "Process step", "Sub process step",
            "Data Title", "Data tag",
            "Théorique value",
        ]
        if KEEP_VISA_VALUE:
            keep_cols.append("Visa value")

        # Keep raw cols too so validation passes
        keep_cols += required_raw

        df = df[[c for c in keep_cols if c in df.columns]]

    return df


def pdf_to_excel(pdf_path: str, xlsx_path: str, compact_output: bool = True) -> None:
    """
    Backward compatible: same behavior as before, but uses pdf_to_dataframe internally.
    """
    df = pdf_to_dataframe(pdf_path, compact_output=compact_output)
    df.to_excel(xlsx_path, index=False)


# =============================================================================
# UI entrypoint: bytes in -> DataFrame out
# =============================================================================

UI_COLUMNS = [
    "PDF name",
    "Data number",
    "Page number",
    "Paragraph number",
    "Process step",
    "Sub process step",
    "Data Title",
    "Data tag",
    "Théorique value",
    "Visa value",
    # keep raw cols for App.py validation
    "Théorique",
    "Réel",
    "Visa",
    "N° observ.",
]

def extract(file_bytes: bytes, file_name: str, product_name: Optional[str] = None) -> pd.DataFrame:
    ext = os.path.splitext(file_name or "")[1].lower()
    if ext != ".pdf":
        raise ValueError("Extractor expects a PDF file.")

    with tempfile.TemporaryDirectory() as td:
        pdf_path = os.path.join(td, "input.pdf")
        with open(pdf_path, "wb") as f:
            f.write(file_bytes)

        df = pdf_to_dataframe(pdf_path, compact_output=True)
        df.columns = [str(c).strip() for c in df.columns]

        # Ensure UI columns exist
        for col in UI_COLUMNS:
            if col not in df.columns:
                df[col] = ""

        # ✅ FIX PDF name column:
        # 1) clear it for all rows
        df["PDF name"] = ""
        # 2) set only the first row to the real uploaded filename
        if len(df) > 0:
            df.loc[df.index[0], "PDF name"] = file_name

        return df[UI_COLUMNS].copy()

# =============================================================================
# Run (optional local testing)
# =============================================================================

if __name__ == "__main__":
    pdf_path = r"C:\Users\dounia.ben.daoud\OneDrive - Servier Monde\AI tool Morpheus\LSI-FORM-3492-24.0-FR (1).pdf"
    out_xlsx = r"C:\Users\dounia.ben.daoud\OneDrive - Servier Monde\AI tool Morpheus\bextratBrtbb.xlsx"

    pdf_to_excel(pdf_path, out_xlsx, compact_output=True)
    print("Saved:", out_xlsx)
