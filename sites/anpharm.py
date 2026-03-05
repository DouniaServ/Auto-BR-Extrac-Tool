# -*- coding: utf-8 -*-

import re
from typing import List, Dict, Tuple, Optional

import pdfplumber
import fitz  # PyMuPDF
import pandas as pd


# ============================================================
# Regex / helpers
# ============================================================
WS_RE = re.compile(r"\s+")
MANY_SPACES_RE = re.compile(r"\s{6,}")
LINEBREAK_RE = re.compile(r"[\r\n]+")
BULLET_RE = re.compile(r"^\s*[-•*]\s+")
BOLD_RE = re.compile(r"(bold|semibold|demibold|black|heavy)", re.I)

PROCESS_RE_DOT = re.compile(r"^\s*(?P<num>\d+)\.\s+\S+")
TWO_DIGIT_SEP = r"(?:\s+|[./-])"

SUBPROC_FULL_LINE_RE = re.compile(
    rf"^\s*(?P<num>\d+)\s+"
    rf"(?P<code>[A-Z]{{1,3}}\d{{2,5}})\s+"
    rf"(?P<a>\d{{2}}){TWO_DIGIT_SEP}(?P<b>\d{{2}})\s+"
    rf"(?P<title>.+?)\s*$"
)

SUBPROC_NUM_CODE_RE = re.compile(
    r"^\s*(?P<num>\d+)\s+(?P<code>[A-Z]{1,3}\d{2,5})\b(?P<rest>.*)$"
)

SUBPROC_AB_ONLY_RE = re.compile(r"^\s*(?P<a>\d{2})\s*(?:[./-]|\s+)\s*(?P<b>\d{2})\s*$")

HEADER_BLOCK_RE = re.compile(
    r"""
    (ANPHARM)|
    (DZIAŁ)|
    (ZAPIS\s+PAKOWANIA)|
    (Batch\s+record)|
    (Opracował)|
    (Akceptował)|
    (Zatwierdził)|
    (Numer\s+serii)|
    (Kod\s+produktu)|
    (Wzór\s+dok)|
    (Strona)
    """,
    re.I | re.X
)

KONTY_RE = re.compile(r"^\s*(kontynuacja|continued)(?:\s*/\s*(kontynuacja|continued))?\s*$", re.I)

DASH_VAL_RE = re.compile(r"^(?P<label>.+?)\s*[−-]\s*[‘']?(?P<val>[^’']+)[’']?\s*$")
WRAPPED_QUOTE_RE = re.compile(r"^[‘'\"](.*?)[’'\"]$")
COLON_VAL_RE = re.compile(r"^(?P<label>.+?)\s*:\s*(?P<val>.+)$")
SPACES_SEP_RE = re.compile(r"^(?P<label>.+?)\s{2,}(?P<val>\S.+)$")

PLACEHOLDER_INT_RE = re.compile(r"^\s*\d{1,5}\s*$")


# ============================================================
# Header detection
# ============================================================
_PL_MAP = str.maketrans({
    "ą": "a", "ć": "c", "ę": "e", "ł": "l", "ń": "n", "ó": "o", "ś": "s", "ż": "z", "ź": "z",
    "Ą": "a", "Ć": "c", "Ę": "e", "Ł": "l", "Ń": "n", "Ó": "o", "Ś": "s", "Ż": "z", "Ź": "z",
})


def norm(s: str) -> str:
    return WS_RE.sub(" ", str(s or "").replace("\u00A0", " ")).strip()


def norm_header(s: str) -> str:
    t = norm(s).translate(_PL_MAP).lower()
    t = re.sub(r"\s+", " ", t).strip()
    return t


HEADER_STRONG = [
    "dzial ap",
    "zapis pakowania",
    "batch record",
    "numer serii",
    "wielkosc serii",
    "kod produktu",
    "wzor dok",
    "strona",
]

HEADER_APPROVAL = ["opracowal", "akceptowal", "zatwierdzil"]
HEADER_PREFIXES = ["numer se", "wielkosc s"]


def is_leaked_page_header_text(text: str) -> bool:
    t = norm_header(text)
    if not t:
        return False

    if any(tok in t for tok in HEADER_APPROVAL):
        return True
    if any(pref in t for pref in HEADER_PREFIXES):
        return True

    strong_hits = sum(1 for tok in HEADER_STRONG if tok in t)
    if strong_hits >= 2:
        return True

    if ("dzial ap" in t) and ("zapis pakowania" in t):
        return True

    return False


def is_leaked_page_header_row(cells: List[str]) -> bool:
    if not cells:
        return False
    joined = " ".join(norm(c) for c in cells if norm(c))
    if is_leaked_page_header_text(joined):
        return True
    for c in cells:
        if is_leaked_page_header_text(c):
            return True
    return False


# ============================================================
# Table header / noise detection
# ============================================================
TABLE_HEADER_TOKENS = {
    "wymaganie", "wykonanie", "wynik", "podpis", "data", "godz.", "godz", "godzina",
    "zgodny", "niezgodny", "z", "nz", "nd",
    "kontynuacja", "lp", "nr", "uwagi",
    "requirement", "execution", "result", "signature", "date", "time",
    "complies", "does not comply", "continued",
}


def _tokenize_simple(s: str) -> List[str]:
    t = norm(s).lower()
    return [x for x in re.split(r"[^\w]+", t) if x]


def is_table_header_row(cells: List[str]) -> bool:
    non_empty = [norm(c) for c in cells if norm(c)]
    if not non_empty:
        return True

    joined = " ".join(non_empty).strip().lower()
    toks = _tokenize_simple(joined)

    header_hits = sum(1 for tok in toks if tok in TABLE_HEADER_TOKENS)
    if header_hits >= 2 and header_hits >= max(2, int(0.6 * len(toks))):
        return True

    if joined in {
        "wymaganie wykonanie data godz. podpis",
        "wymaganie wykonanie data godz podpis",
        "wynik podpis",
        "wynik z nz podpis",
        "z nz",
        "z nz nd",
        "data godz. podpis",
        "data godz podpis",
    }:
        return True

    # avoid false positives like "Data produkcji / Manufacturing date"
    if len(joined) <= 35 and len(toks) <= 6:
        header_hits2 = sum(1 for tok in toks if tok in TABLE_HEADER_TOKENS)
        if header_hits2 >= 2:
            return True

    return False


def looks_like_instruction_row(text: str) -> bool:
    t = norm(text).lower()
    starters = ["zaznacz", "wypełnij", "wpisz", "potwierdź", "sprawdź", "fill", "mark", "tick", "enter"]
    return any(t.startswith(s) for s in starters)


def unquote_if_wrapped(s: str) -> str:
    t = norm(s)
    m = WRAPPED_QUOTE_RE.match(t)
    return norm(m.group(1)) if m else t


def split_inline_dash_value(label: str) -> Tuple[str, str]:
    t = norm(label)
    m = DASH_VAL_RE.match(t)
    if not m:
        return t, ""
    return norm(m.group("label")), norm(m.group("val"))


def looks_like_rotated_garbage(s: str) -> bool:
    t = norm(s)
    if len(t) < 25:
        return False
    letters = sum(ch.isalpha() for ch in t)
    ratio = letters / max(1, len(t))
    return ratio < 0.45


def is_top_banner(process: str) -> bool:
    t = (process or "").lower()
    return ("zapis pakowania" in t) or ("batch record" in t) or ("dział" in t)


def parse_label_value_from_line(line: str) -> Tuple[str, str]:
    t = unquote_if_wrapped(line)

    lbl, v = split_inline_dash_value(t)
    if v:
        return norm(lbl), unquote_if_wrapped(v)

    m = COLON_VAL_RE.match(t)
    if m:
        return norm(m.group("label")), unquote_if_wrapped(norm(m.group("val")))

    m = SPACES_SEP_RE.match(t)
    if m and len(m.group("label")) <= 180:
        return norm(m.group("label")), unquote_if_wrapped(norm(m.group("val")))

    return t, ""


# ============================================================
# ✅ ROBUST PAGE-1 WORDS-BASED ANCHOR EXTRACTION
# ============================================================
def extract_page1_anchor_fields(pdf_path: str) -> List[Dict]:
    """
    Extract specific label/value fields from page 1 using get_text("words"),
    grouped by (block,line). This is much more stable than y-grouping.
    """
    doc = fitz.open(pdf_path)
    page = doc[0]

    # (x0, y0, x1, y1, word, block_no, line_no, word_no)
    words = page.get_text("words") or []
    doc.close()

    # group words by (block, line)
    lines_map: Dict[Tuple[int, int], List[tuple]] = {}
    for w in words:
        if len(w) < 8:
            continue
        x0, y0, x1, y1, text, bno, lno, wno = w
        text = norm(text)
        if not text:
            continue
        lines_map.setdefault((int(bno), int(lno)), []).append((float(x0), float(x1), float(y0), text))

    # build ordered lines
    ordered = []
    for (bno, lno), items in lines_map.items():
        items.sort(key=lambda t: t[0])
        line_text = norm(" ".join(t[3] for t in items))
        if not line_text:
            continue
        y = min(t[2] for t in items)
        ordered.append((y, bno, lno, line_text))
    ordered.sort(key=lambda t: (t[0], t[1], t[2]))

    # helper: find next non-empty line text
    def next_line_text(idx: int) -> str:
        for j in range(idx + 1, min(idx + 4, len(ordered))):
            t = ordered[j][3]
            if t and not is_leaked_page_header_text(t):
                return t
        return ""

    anchor_patterns = [
        re.compile(r"nazwa\s+produktu\s+luzem", re.I),
        re.compile(r"bulk\s+product\s+name", re.I),
    ]

    out: List[Dict] = []
    seen_keys = set()

    for i, (_y, _b, _l, line_text) in enumerate(ordered):
        t = line_text
        if not t:
            continue
        if is_leaked_page_header_text(t):
            continue
        if looks_like_instruction_row(t) and len(t) < 200:
            continue

        if not any(p.search(t) for p in anchor_patterns):
            continue

        label, val = parse_label_value_from_line(t)

        if (not val) and label.rstrip().endswith(":"):
            nxt = next_line_text(i)
            if nxt:
                if WRAPPED_QUOTE_RE.match(norm(nxt)) or ("Prestarium" in nxt) or (len(nxt) <= 80):
                    val = unquote_if_wrapped(nxt)

        if not val:
            m = re.search(r"[:−-]\s*(.+)$", t)
            if m and m.group(1):
                val = unquote_if_wrapped(m.group(1).strip())
                label = norm(t[:m.start(1)].rstrip())

        label = norm(label)
        val = norm(val)

        if not label:
            continue
        if is_leaked_page_header_text(label) or is_leaked_page_header_text(val):
            continue

        key = (1, label.lower(), val.lower())
        if key in seen_keys:
            continue
        seen_keys.add(key)

        out.append({
            "Page number": 1,
            "Paragraph number": 1,
            "Process step": "",
            "Sub Process step": "",
            "Data title": label,
            "Value": val,
        })

    return out


# ============================================================
# ✅ FIXED: Text line grouping (PyMuPDF) for pages 2+
# ============================================================
def page_lines_grouped(
    page: fitz.Page,
    y_min_ratio: float,
    y_max_ratio: float,
    y_tol: float = 2.6,
) -> List[Dict]:
    ph = page.rect.height
    pw = page.rect.width
    y_min = ph * y_min_ratio
    y_max = ph * y_max_ratio

    spans = []
    d = page.get_text("dict")
    for b in d.get("blocks", []):
        for ln in b.get("lines", []):
            for sp in ln.get("spans", []):
                txt = norm(sp.get("text", ""))
                if not txt:
                    continue
                x0, y0, x1, y1 = sp.get("bbox", (0, 0, 0, 0))
                if y0 < y_min or y0 > y_max:
                    continue
                font = sp.get("font", "") or ""
                bold = bool(BOLD_RE.search(font)) or bool(sp.get("flags", 0) & 2)
                size = float(sp.get("size", 0) or 0)
                spans.append((float(y0), float(x0), float(x1), txt, bold, size))

    spans.sort(key=lambda t: (t[0], t[1]))

    rows: List[List[tuple]] = []
    for y0, x0, x1, txt, bold, size in spans:
        placed = False
        for r in rows:
            if abs(y0 - r[0][0]) <= y_tol:
                r.append((y0, x0, x1, txt, bold, size))
                placed = True
                break
        if not placed:
            rows.append([(y0, x0, x1, txt, bold, size)])

    gap_tol = max(80.0, 0.14 * pw)

    lines: List[Dict] = []
    for r in rows:
        r.sort(key=lambda t: t[1])

        seg: List[tuple] = []
        prev_x1 = None

        def flush_segment(segment: List[tuple]):
            if not segment:
                return
            text = norm(" ".join(t[3] for t in segment))
            if not text:
                return
            lines.append({
                "text": text,
                "any_bold": any(t[4] for t in segment),
                "y": segment[0][0],
                "x0": segment[0][1],
                "max_size": max((t[5] for t in segment), default=0.0),
            })

        for item in r:
            y0, x0, x1, txt, bold, size = item
            if prev_x1 is not None and (x0 - prev_x1) > gap_tol:
                flush_segment(seg)
                seg = []
                prev_x1 = None

            seg.append(item)
            prev_x1 = x1

        flush_segment(seg)

    return lines


# ============================================================
# Process step detection (unchanged)
# ============================================================
def _upper_ratio(s: str) -> float:
    letters = [ch for ch in s if ch.isalpha()]
    if not letters:
        return 0.0
    upp = sum(ch.isupper() for ch in letters)
    return upp / max(1, len(letters))


def _is_all_caps_heading(s: str) -> bool:
    t = norm(s)
    if len(t) < 8:
        return False
    if HEADER_BLOCK_RE.search(t):
        return False
    return _upper_ratio(t) >= 0.85 and len(t.split()) >= 2


def detect_process_step_from_band(
    lines: List[Dict],
    prev_process_num: Optional[int],
    prev_process_size: float,
) -> Tuple[str, Optional[int], float]:
    dotted = []
    for ln in lines:
        t = ln["text"]
        if HEADER_BLOCK_RE.search(t):
            continue
        m = PROCESS_RE_DOT.match(t)
        if m:
            dotted.append((ln["max_size"], ln["y"], t, int(m.group("num"))))

    caps = []
    for ln in lines:
        t = ln["text"]
        if _is_all_caps_heading(t):
            caps.append((ln["max_size"], ln["y"], t))

    best_text = ""
    best_num: Optional[int] = None
    best_size = 0.0

    if dotted:
        dotted.sort(key=lambda x: (-x[0], x[1]))
        best_size, _, best_text, best_num = dotted[0]

        if prev_process_num is not None and best_num is not None and best_num < prev_process_num:
            if best_size + 0.2 < prev_process_size:
                best_text, best_num, best_size = "", None, 0.0

        if best_text and prev_process_size > 0 and best_size + 0.2 < prev_process_size:
            best_text, best_num, best_size = "", None, 0.0

    if not best_text and caps:
        caps.sort(key=lambda x: (-x[0], x[1]))
        best_size, _, best_text = caps[0]
        best_num = None
        if prev_process_size > 0 and best_size + 0.2 < prev_process_size:
            best_text, best_size = "", 0.0

    return best_text, best_num, best_size


# ============================================================
# Subprocess detection + cleaning (unchanged)
# ============================================================
SUBPROC_CUT_TOKENS_RE = re.compile(
    r"""
    \b(
        Wynik|Podpis|Wymaganie|Wykonanie|Data|Godz\.?|Godzina|
        Signature|Requirement|Execution|Result
    )\b
    """,
    re.I | re.X
)

SUBPROC_INSTR_CUT_RE = re.compile(
    r"\b(W\s+przypadku|Wypełnij|Zaznacz|Wpisz|Potwierdź|Sprawdź)\b",
    re.I
)


def _clean_rotated_append(t: str) -> str:
    t = norm(t)

    if " Butelkowanie - " in t:
        t = t.split(" Butelkowanie - ")[0].strip()

    m_instr = SUBPROC_INSTR_CUT_RE.search(t)
    if m_instr and m_instr.start() >= 10:
        t = t[:m_instr.start()].strip()

    m_tok = SUBPROC_CUT_TOKENS_RE.search(t)
    if m_tok and m_tok.start() >= 12:
        t = t[:m_tok.start()].strip()

    return t


def detect_subprocess_from_header_band(lines: List[Dict]) -> str:
    for ln in lines:
        t = _clean_rotated_append(ln["text"])
        if HEADER_BLOCK_RE.search(t):
            continue
        m = SUBPROC_FULL_LINE_RE.match(t)
        if m:
            return _clean_rotated_append(t)

    for i, ln in enumerate(lines):
        t0 = _clean_rotated_append(ln["text"])
        if HEADER_BLOCK_RE.search(t0):
            continue

        m0 = SUBPROC_NUM_CODE_RE.match(t0)
        if not m0:
            continue

        num = m0.group("num")
        code = m0.group("code")
        rest = norm(m0.group("rest") or "")

        try:
            if int(num) > 99:
                continue
        except Exception:
            continue

        a = b = ""
        title_parts: List[str] = []

        m_ab_inline = re.search(r"\b(\d{2})\s*(?:[./-]|\s+)\s*(\d{2})\b", rest)
        if m_ab_inline:
            a, b = m_ab_inline.group(1), m_ab_inline.group(2)
            rest_clean = norm(re.sub(r"\b\d{2}\s*(?:[./-]|\s+)\s*\d{2}\b", "", rest))
            if rest_clean:
                title_parts.append(rest_clean)
        else:
            if rest:
                title_parts.append(rest)

        for j in range(i + 1, min(i + 7, len(lines))):
            t = _clean_rotated_append(lines[j]["text"])
            if not t:
                continue
            if HEADER_BLOCK_RE.search(t):
                continue
            if KONTY_RE.match(t):
                continue
            if PROCESS_RE_DOT.match(t) or SUBPROC_FULL_LINE_RE.match(t):
                break
            if SUBPROC_NUM_CODE_RE.match(t) and j != i:
                break

            m_ab = SUBPROC_AB_ONLY_RE.match(t)
            if m_ab and not (a and b):
                a, b = m_ab.group("a"), m_ab.group("b")
                continue

            title_parts.append(t)
            if len(title_parts) >= 2:
                break

        prefix = f"{num} {code}"
        if a and b:
            prefix = f"{prefix} {a} {b}"

        title = norm(" ".join([p for p in title_parts if p]))
        return _clean_rotated_append(norm(prefix + (" " + title if title else "")))

    return ""


# ============================================================
# Page 1 — tables (your existing logic)
# ============================================================
def extract_page1_tables(pdf_path: str) -> List[Dict]:
    records: List[Dict] = []

    def _looks_like_value(s: str) -> bool:
        t = norm(s)
        if not t:
            return False
        if WRAPPED_QUOTE_RE.match(t):
            return True
        if re.fullmatch(r"[0-9.,/]+", t):
            return True
        if len(t) <= 40 and any(ch.isdigit() for ch in t):
            return True
        if len(t) <= 140 and not any(x in t.lower() for x in ["wymaganie", "wykonanie", "wynik", "podpis"]):
            return True
        return False

    def _looks_like_label(s: str) -> bool:
        t = norm(s)
        if not t:
            return False
        if HEADER_BLOCK_RE.search(t):
            return False
        return ("/" in t) or t.rstrip().endswith(":") or (len(t) >= 8)

    with pdfplumber.open(pdf_path) as pdf:
        page = pdf.pages[0]
        tables = page.extract_tables() or []

        paragraph_table_no = 0

        for table in tables:
            if not table or not table[0]:
                continue

            header_cells = [norm(c) for c in table[0] if c is not None]
            if not any(header_cells):
                continue

            candidate_process = header_cells[0] if header_cells else ""

            if is_top_banner(candidate_process) or HEADER_BLOCK_RE.search(candidate_process):
                continue

            paragraph_table_no += 1
            process = norm(candidate_process)

            header_non_empty = [unquote_if_wrapped(c) for c in header_cells if norm(c)]
            first_row_emitted = False

            if (
                len(header_non_empty) >= 2
                and _looks_like_label(header_non_empty[0])
                and _looks_like_value(header_non_empty[-1])
                and not is_table_header_row(header_non_empty)
            ):
                left = header_non_empty[0]
                right = header_non_empty[-1]

                lbl_parsed, val_parsed = parse_label_value_from_line(left)
                if val_parsed and not right:
                    left, right = lbl_parsed, val_parsed

                records.append({
                    "Page number": 1,
                    "Paragraph number": paragraph_table_no,
                    "Process step": process,
                    "Sub Process step": "",
                    "Data title": left,
                    "Value": right,
                })
                first_row_emitted = True

            elif len(header_non_empty) == 1:
                lbl_parsed, val_parsed = parse_label_value_from_line(header_non_empty[0])
                if val_parsed and _looks_like_label(lbl_parsed) and _looks_like_value(val_parsed):
                    records.append({
                        "Page number": 1,
                        "Paragraph number": paragraph_table_no,
                        "Process step": process,
                        "Sub Process step": "",
                        "Data title": lbl_parsed,
                        "Value": val_parsed,
                    })
                    first_row_emitted = True

            if not first_row_emitted:
                extras = [unquote_if_wrapped(c) for c in header_cells[1:] if norm(c)]
                for ex in extras:
                    if not ex:
                        continue
                    if looks_like_instruction_row(ex):
                        continue
                    if HEADER_BLOCK_RE.search(ex):
                        continue
                    records.append({
                        "Page number": 1,
                        "Paragraph number": paragraph_table_no,
                        "Process step": process,
                        "Sub Process step": "",
                        "Data title": ex,
                        "Value": ex,
                    })

            for row in table[1:]:
                cells = [norm(c) for c in (row or [])]
                if not any(cells):
                    continue

                non_empty = [unquote_if_wrapped(c) for c in cells if norm(c)]
                if not non_empty:
                    continue

                if is_table_header_row(non_empty):
                    continue

                joined = norm(" ".join(non_empty))
                if looks_like_rotated_garbage(joined):
                    continue
                if looks_like_instruction_row(joined) and len(joined) < 160:
                    continue

                if len(non_empty) == 1:
                    label = non_empty[0]
                    if not label or HEADER_BLOCK_RE.search(label):
                        continue
                    records.append({
                        "Page number": 1,
                        "Paragraph number": paragraph_table_no,
                        "Process step": process,
                        "Sub Process step": "",
                        "Data title": label,
                        "Value": "",
                    })
                    continue

                left = non_empty[0]
                right = non_empty[-1] if len(non_empty) > 1 else ""

                if not left or HEADER_BLOCK_RE.search(left):
                    continue

                lbl2, inline_val = split_inline_dash_value(left)
                if inline_val:
                    left = lbl2
                    right = inline_val

                records.append({
                    "Page number": 1,
                    "Paragraph number": paragraph_table_no,
                    "Process step": process,
                    "Sub Process step": "",
                    "Data title": left,
                    "Value": right,
                })

    return records


# ============================================================
# Tables on pages 2+
# ============================================================
def build_label_value_from_cells(non_empty: List[str]) -> Tuple[str, str]:
    if not non_empty:
        return "", ""
    if len(non_empty) == 1:
        return non_empty[0], ""

    left_parts = non_empty[:-1]
    right = non_empty[-1]

    joined_left = norm(" ".join(left_parts))
    left = joined_left if len(joined_left) <= 220 else left_parts[0]

    left = unquote_if_wrapped(left)
    right = unquote_if_wrapped(right)

    lbl2, inline_val = split_inline_dash_value(left)
    if inline_val:
        return lbl2, inline_val

    return left, right


def extract_tables_with_subprocess_from_row(
    pdf_path: str,
    page_index0: int,
    current_process: str,
    current_sub: str
) -> List[Dict]:
    rows_out: List[Dict] = []

    with pdfplumber.open(pdf_path) as pdf:
        page = pdf.pages[page_index0]
        tables = page.extract_tables() or []

        for table in tables:
            for row in table:
                cells = [norm(c) for c in (row or [])]
                if not any(cells):
                    continue

                if is_leaked_page_header_row(cells):
                    continue

                if is_table_header_row(cells):
                    continue

                non_empty = [unquote_if_wrapped(c) for c in cells if norm(c)]
                if not non_empty or is_table_header_row(non_empty):
                    continue

                if is_leaked_page_header_text(" ".join(non_empty)):
                    continue

                row_joined = norm(" ".join(non_empty))
                if looks_like_instruction_row(row_joined) and len(row_joined) < 160:
                    continue
                if looks_like_rotated_garbage(row_joined):
                    continue

                label, value = build_label_value_from_cells(non_empty)
                if not label:
                    continue

                if is_leaked_page_header_text(label) or is_leaked_page_header_text(value):
                    continue

                if looks_like_instruction_row(label) and not value:
                    continue

                rows_out.append({
                    "Page number": page_index0 + 1,
                    "Paragraph number": None,
                    "Process step": current_process,
                    "Sub Process step": current_sub,
                    "Data title": label,
                    "Value": value,
                })

    return rows_out


# ============================================================
# Pages 2+ extraction
# ============================================================
def extract_pages_2plus_target(pdf_path: str, start_page: int = 2) -> List[Dict]:
    doc = fitz.open(pdf_path)
    out: List[Dict] = []

    current_process = ""
    current_process_num: Optional[int] = None
    current_process_size: float = 0.0
    current_sub = ""

    for pno0 in range(start_page - 1, len(doc)):
        page = doc[pno0]
        page_no = pno0 + 1

        header_lines = page_lines_grouped(page, y_min_ratio=0.00, y_max_ratio=0.35, y_tol=2.6)
        body_lines = page_lines_grouped(page, y_min_ratio=0.10, y_max_ratio=0.94, y_tol=2.6)

        proc, pnum, psize = detect_process_step_from_band(
            header_lines,
            prev_process_num=current_process_num,
            prev_process_size=current_process_size
        )
        if proc:
            current_process = proc
            current_process_num = pnum
            current_process_size = psize
            current_sub = ""

        sub = detect_subprocess_from_header_band(header_lines)
        if sub:
            current_sub = sub

        paragraph_no = 1

        table_rows = extract_tables_with_subprocess_from_row(pdf_path, pno0, current_process, current_sub)

        seen = set()
        cleaned_table_rows = []
        for r in table_rows:
            key = (r["Process step"], r["Sub Process step"], r["Data title"], r["Value"])
            if key in seen:
                continue
            seen.add(key)
            cleaned_table_rows.append(r)

        for r in cleaned_table_rows:
            r["Paragraph number"] = paragraph_no
            out.append(r)

        existing_titles = set(r["Data title"] for r in cleaned_table_rows)

        i = 0
        while i < len(body_lines):
            t = body_lines[i]["text"]
            i += 1
            if not t:
                continue

            if is_leaked_page_header_text(t):
                continue
            if t == current_process or t == current_sub:
                continue
            if looks_like_rotated_garbage(t):
                continue
            if re.match(r"^\s*\d+\s*/\s*\d+\s*$", t):
                continue
            if looks_like_instruction_row(t) and len(t) < 160:
                continue

            if t in existing_titles:
                _lbl_tmp, _val_tmp = parse_label_value_from_line(t)
                if not _val_tmp:
                    continue

            label, val = parse_label_value_from_line(t)
            if not label:
                continue

            if (not val) and label.rstrip().endswith(":") and i < len(body_lines):
                nxt = body_lines[i]["text"]
                if nxt:
                    if not is_leaked_page_header_text(nxt) and not looks_like_instruction_row(nxt):
                        if not re.match(r"^\s*\d+\s*/\s*\d+\s*$", nxt) and not MANY_SPACES_RE.search(nxt):
                            val = unquote_if_wrapped(nxt)
                            i += 1

            if is_leaked_page_header_text(label) or is_leaked_page_header_text(val):
                continue

            out.append({
                "Page number": page_no,
                "Paragraph number": paragraph_no,
                "Process step": current_process,
                "Sub Process step": current_sub,
                "Data title": label,
                "Value": val,
            })

    doc.close()
    return out


# ============================================================
# Tag + review
# ============================================================
def compute_data_tag_row(data_title: str, value: str, process_step: str) -> str:
    t = norm(data_title)
    v = norm(value)
    if not t and not v:
        return "empty"
    if t and PLACEHOLDER_INT_RE.match(t) and not v:
        return "placeholder_row"
    if t and v:
        return "field"
    if t and not v:
        if BULLET_RE.match(t):
            return "bullet_item"
        return "title"
    return "title"


def add_needs_review_and_tag(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()

    for c in ["Page number", "Paragraph number", "Process step", "Sub Process step", "Data title", "Value"]:
        if c not in out.columns:
            out[c] = ""

    title = out["Data title"].fillna("").astype(str)
    value = out["Value"].fillna("").astype(str)
    proc = out["Process step"].fillna("").astype(str)

    out["Data tag"] = [compute_data_tag_row(t, v, p) for t, v, p in zip(title, value, proc)]

    out["Needs review"] = "No"
    out["Review reasons"] = "-"

    mask_noise = title.str.contains(LINEBREAK_RE) | title.str.contains(MANY_SPACES_RE)
    out.loc[mask_noise, "Needs review"] = "Yes"
    out.loc[mask_noise, "Review reasons"] = "layout_noise"

    return out


# ============================================================
# ✅ Post-cleaning: skip punctuation-only rows + merge 1-word fragments
# (does NOT change extraction logic; only cleans extracted output)
# ============================================================
PUNCT_FILLER_ONLY_RE = re.compile(
    r"""^\s*(
        [\u2022\u25CF•●\-\–\—\*_=\.·,;:|/\\]+ |   # bullets/punct/fillers
        [‘’'"`]+                                  # quotes only
    )\s*$""",
    re.X
)

UNIT_ONLY_TOKENS = {
    "kg", "g", "mg", "µg", "ug", "ml", "l", "cl", "dl",
    "%", "ppm"
}


def _is_punct_or_filler_only(title: str) -> bool:
    t = norm(title)
    if not t:
        return True
    if PUNCT_FILLER_ONLY_RE.match(t):
        return True
    if re.fullmatch(r"[_\-\s]{3,}", t):
        return True
    if re.fullmatch(r"\*{2,}", t):
        return True
    return False


def _is_one_word_fragment(title: str) -> bool:
    t = norm(title)
    if not t:
        return False
    toks = t.split()
    if len(toks) != 1:
        return False
    if re.fullmatch(r"\d+([.,]\d+)?", toks[0]):
        return False
    return True


def _should_merge_into_previous(row_title: str, row_value: str) -> bool:
    t = norm(row_title)
    v = norm(row_value)
    if not t and not v:
        return False

    if t.lower() in UNIT_ONLY_TOKENS:
        return True

    if _is_punct_or_filler_only(t):
        return bool(v)

    if _is_one_word_fragment(t):
        return True

    return False


def _same_context(prev: Dict, cur: Dict) -> bool:
    keys = ["Page number", "Paragraph number", "Process step", "Sub Process step"]
    return all(norm(prev.get(k, "")) == norm(cur.get(k, "")) for k in keys)


def clean_skip_and_merge_fragments(rows: List[Dict]) -> List[Dict]:
    """
    - Drops rows where Data title is punctuation/filler only (●, ***, _____, quotes-only, etc.)
    - Merges 1-word / unit-only / fragment rows into previous row's Data title
      (and includes current Value too, if present), then drops the fragment row.
    """
    cleaned: List[Dict] = []

    for r in rows:
        title = norm(r.get("Data title", ""))
        value = norm(r.get("Value", ""))

        # Drop pure filler rows (punct only) if no value
        if _is_punct_or_filler_only(title) and not value:
            continue

        # Merge fragment rows into previous row's Data title (same context)
        if cleaned and _should_merge_into_previous(title, value) and _same_context(cleaned[-1], r):
            prev = cleaned[-1]
            prev_title = norm(prev.get("Data title", ""))

            fragment = title
            if value:
                fragment = norm(f"{title} {value}") if title else value

            if fragment:
                prev["Data title"] = norm(f"{prev_title} {fragment}") if prev_title else fragment

            # keep previous Value unchanged (integrate fragment into Data title)
            continue

        r["Data title"] = title
        r["Value"] = value
        cleaned.append(r)

    return cleaned


# ============================================================
# Main export (✅ includes robust page1 anchor fallback + cleaning)
# ============================================================
def pdf_to_excel_structured(pdf_path: str, out_xlsx: str, include_review_cols: bool = True):
    page1_tables = extract_page1_tables(pdf_path)
    page1_anchor = extract_page1_anchor_fields(pdf_path)

    # merge + dedup page1
    seen = set()
    page1 = []
    for r in (page1_tables + page1_anchor):
        key = (r.get("Page number"), r.get("Data title", "").strip().lower(), r.get("Value", "").strip().lower())
        if key in seen:
            continue
        seen.add(key)
        page1.append(r)

    p2plus = extract_pages_2plus_target(pdf_path, start_page=2)

    # ✅ NEW: clean output rows (skip punctuation-only + merge fragments)
    all_rows = clean_skip_and_merge_fragments(page1 + p2plus)

    if not all_rows:
        raise RuntimeError("No rows extracted. If the PDF is scanned, OCR is required.")

    df = pd.DataFrame(all_rows).fillna("")
    df.insert(0, "Data number", range(1, len(df) + 1))

    cols = [
        "Data number",
        "Page number",
        "Paragraph number",
        "Process step",
        "Sub Process step",
        "Data title",
        "Value",
    ]
    for c in cols:
        if c not in df.columns:
            df[c] = ""
    df = df[cols]

    if include_review_cols:
        df = add_needs_review_and_tag(df)
        df = df[cols + ["Data tag", "Needs review", "Review reasons"]]

    df.to_excel(out_xlsx, index=False)
    print(f"Saved: {out_xlsx}")
    print(f"Rows: {len(df)}")


# ============================================================
# UI extract wrapper (✅ same extraction, plus cleaning)
# ============================================================
import os
import tempfile
from typing import Optional


def extract(file_bytes: bytes, file_name: str, product_name: Optional[str] = None) -> pd.DataFrame:
    ext = os.path.splitext(file_name or "")[1].lower()
    if ext != ".pdf":
        raise ValueError("Extractor expects a PDF file.")

    with tempfile.TemporaryDirectory() as td:
        pdf_path = os.path.join(td, "input.pdf")
        with open(pdf_path, "wb") as f:
            f.write(file_bytes)

        # --- SAME extraction as your original code ---
        page1_tables = extract_page1_tables(pdf_path)
        page1_anchor = extract_page1_anchor_fields(pdf_path)

        seen = set()
        page1 = []
        for r in (page1_tables + page1_anchor):
            key = (
                r.get("Page number"),
                (r.get("Data title", "") or "").strip().lower(),
                (r.get("Value", "") or "").strip().lower(),
            )
            if key in seen:
                continue
            seen.add(key)
            page1.append(r)

        p2plus = extract_pages_2plus_target(pdf_path, start_page=2)

        # ✅ NEW: clean output rows (skip punctuation-only + merge fragments)
        all_rows = clean_skip_and_merge_fragments(page1 + p2plus)

        if not all_rows:
            raise RuntimeError("No rows extracted. If the PDF is scanned, OCR is required.")

        df = pd.DataFrame(all_rows).fillna("")
        df.insert(0, "Data number", range(1, len(df) + 1))
        df = add_needs_review_and_tag(df)

        # --- UI adapter (fix column names exactly as UI expects) ---
        ui_df = df.copy()

        # PDF name column (first row only)
        ui_df["PDF name"] = ""
        if len(ui_df) > 0:
            ui_df.at[0, "PDF name"] = file_name

        # ✅ fix exact expected names
        ui_df = ui_df.rename(columns={
            "Sub Process step": "Sub process step",  # UI expects this casing
            "Data title": "Data Title",              # UI expects capital T
        })

        # ✅ guarantee required columns exist (avoid future ValueError)
        required = [
            "PDF name",
            "Data number",
            "Page number",
            "Paragraph number",
            "Process step",
            "Sub process step",
            "Data Title",
            "Value",
            "Data tag",
        ]
        for c in required:
            if c not in ui_df.columns:
                ui_df[c] = ""

        # ✅ If Sub process step is empty => "--"
        ui_df["Sub process step"] = (
            ui_df["Sub process step"]
            .fillna("")
            .astype(str)
            .str.strip()
        )
        ui_df.loc[ui_df["Sub process step"] == "", "Sub process step"] = "--"

        ui_df["Data Title"] = ui_df["Data Title"].fillna("").astype(str)
        ui_df["Value"] = ui_df["Value"].fillna("").astype(str)

        return ui_df[required].copy()