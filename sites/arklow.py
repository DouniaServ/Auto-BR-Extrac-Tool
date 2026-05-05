import os
import re
import tempfile
import fitz
import pandas as pd
import unicodedata
from dataclasses import dataclass
from typing import List, Dict, Tuple, Optional, Set
from collections import defaultdict

# ============================================================
# UI OUTPUT COLUMNS
# ============================================================
COLUMNS = [
    "PDF name",          # MUST be the input PDF file name (only first row)
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

# ============================================================
# Normalization helpers
# ============================================================
ZERO_WIDTH_RE = re.compile(r"[\u200B-\u200F\u202A-\u202E\u2060\uFEFF]")
NBSP_RE = re.compile(r"[\u00A0\u202F]")

def norm(s: str) -> str:
    s = s or ""
    s = unicodedata.normalize("NFKC", s)
    s = NBSP_RE.sub(" ", s)
    s = ZERO_WIDTH_RE.sub("", s)
    return re.sub(r"\s+", " ", s).strip()

def tag_data_safe(tag: str) -> str:
    t = (tag or "").strip().upper()
    return t if t in ALLOWED_TAG_DATA else "TEXT"

# ============================================================
# Skip patterns: headers / batch size / DocuSign / etc
# IMPORTANT: DO NOT drop "PP.M.MAN.PROC01" etc (routing table templates)
# ============================================================
HEADER_DROP_REGEX = [
    r"^docusign envelope id:",
    r"^page\s+\d+\s+of\s+\d+",
    r"^servier\s*\(ireland\)\s*industries(?:\s+ltd\.?)?$",
    r"^batch\s*record\b.*$",
    r"^document\s*name\b.*$",
    r"^batch\s*number\b.*$",
    r"^batch\s*size\b.*$",
    r"^document\s*number\b.*$",
    r"^supercedes\b.*$",
    r"^Document\s*review\s*date\b[:\-]?\s*.*$",
    r"^new\s+document$",

    # ONLY drop edition header line patterns, NOT all PP.M.MAN.*
    r"^pp\.m\.man\.\d{3}\.\d{2}\.\d+\.ed\d+(?:\.\d+)?[a-z]?$",
    r"^pp\.m\.man\.[a-z0-9.]+\.ed\d+(?:\.\d+)?[a-z]?$",

    r"^\d{2}/\d{4}$",
    r"^\d{2}/\d{2}/\d{4}$",
]
HEADER_DROP_PATTERNS = [re.compile(p, re.IGNORECASE) for p in HEADER_DROP_REGEX]

HEADER_LABEL_ONLY_RE = re.compile(
    r"""(?ix)^\s*(
        document|
        document\s*review\s*date|
        review\s*date|
        review|
        title|
        signature|
        date
    )\s*$"""
)

def is_header_label_only(text: str) -> bool:
    return bool(HEADER_LABEL_ONLY_RE.match(norm(text)))

BATCH_SIZE_VALUE_RE = re.compile(r"^\d+(\.\d+)?\s*kg\s*/\s*[\d,]+\s*tablets$", re.IGNORECASE)

# ============================================================
# DocuSign certificate pages
# ============================================================
DOCUSIGN_CERT_TITLE_RE = re.compile(r"\bcertificate\s+of\s+completion\b", re.IGNORECASE)
DOCUSIGN_MARKERS = [
    re.compile(r"\bsigner\s+events\b", re.IGNORECASE),
    re.compile(r"\benvelope\s+summary\s+events\b", re.IGNORECASE),
    re.compile(r"\bsecurity\s+checked\b", re.IGNORECASE),
    re.compile(r"\belectronic\s+record\s+and\s+signature\s+disclosure\b", re.IGNORECASE),
    re.compile(r"\busing\s+ip\s+address\b", re.IGNORECASE),
    re.compile(r"\bviewed:\s*\d{2}-\d{2}-\d{2}\b", re.IGNORECASE),
    re.compile(r"\bsent:\s*\d{2}-\d{2}-\d{2}\b", re.IGNORECASE),
    re.compile(r"\bsigned:\s*\d{2}-\d{2}-\d{2}\b", re.IGNORECASE),
    re.compile(r"\bdocusign\s+envelope\s+id\b", re.IGNORECASE),
]

def page_is_docusign_certificate(page_text: str) -> bool:
    if not page_text:
        return False
    top_lines = "\n".join((page_text.splitlines() or [])[:30])
    has_title = bool(DOCUSIGN_CERT_TITLE_RE.search(top_lines))
    hits = sum(1 for rx in DOCUSIGN_MARKERS if rx.search(page_text))
    return (has_title and hits >= 1) or (hits >= 3)

# ============================================================
# Stage + section patterns
# ============================================================
STAGE_RE = re.compile(r"^\s*stage\s*0?(\d+)\s*(?::\s*)?(.+?)\s*$", re.IGNORECASE)

SECTION_RE = re.compile(
    r"^\s*(raw materials required|lubricants|dispensing equipment|container required|coating materials|theoretical total weight)\s*$",
    re.IGNORECASE,
)

CAPS_SECTION_RE = re.compile(r"^[A-Z][A-Z0-9 /&\-]{3,}$")
CAPS_SECTION_BLOCKLIST = [
    "BATCH RECORD",
    "DOCUMENT NAME",
    "DOCUMENT NUMBER",
    "DOCUMENT REVIEW DATE",
    "HISTORY OF CHANGE",
    "TRAINING RECORD",
    "PAGE",
]

NO_STAGE_SECTION_RE = re.compile(
    r"""(?ix)^\s*(
        .*checks\b |
        process\s+instructions\b |
        compression\s+calculations\b |
        observation\s+page\b |
        comments\b |
        checklist\b |
        tool\s+scanning\s+review\b |
        production\s+review\b |
        quality\s+review\b |
        manual\s+cleaning\b.* |
        cleaning\b.* |
        reconciliation\b |
        in[-\s]?process\s+control.* |
        ipc.* |
        equipment.* |
        start[-\s]?up.* |
        shutdown.* |
        changeover.* |
        line\s+clearance.* |
        sampling.* |
        testing.* |
        deviations?.*
    )\s*$"""
)

def norm_sub_process_step(stage_name: str) -> str:
    return (stage_name or "").strip()

# ============================================================
# Noise filtering + glyph normalization
# ============================================================
NOISE_GLYPHS = {"", "•"}
PUA_RE = re.compile(r"[\uE000-\uF8FF]")
PURE_SYMBOL_RE = re.compile(r"^[^A-Za-z0-9]+$")
MAX_NOISE_SYMBOL_LEN = 12

MATH_GLYPH_MAP = {
    "": ")",
    "": "(",
    "": "(",
    "": ")",
    "": "(",
    "": ")",
}

def normalize_math_glyphs(s: str) -> str:
    s = s or ""
    s = unicodedata.normalize("NFKD", s)
    for k, v in MATH_GLYPH_MAP.items():
        s = s.replace(k, v)
    return norm(s)

REPL_CHAR = "\uFFFD"
PLUS_MINUS_PUA = {"", "", ""}
DEGREE_PUA = {"", "", ""}

def normalize_special_glyphs(s: str) -> str:
    s = s or ""
    s = unicodedata.normalize("NFKD", s)

    for g in PLUS_MINUS_PUA:
        s = s.replace(g, "±")
    for g in DEGREE_PUA:
        s = s.replace(g, "°")

    s = s.replace("Â±", "±").replace("Â°", "°")
    s = re.sub(r"(?<=\d)\s*" + re.escape(REPL_CHAR) + r"\s*(?=\d)", " ± ", s)
    s = re.sub(r"(?<=\d)" + re.escape(REPL_CHAR) + r"(?=\d)", "±", s)
    s = re.sub(r"\s*°\s*C\b", " °C", s, flags=re.IGNORECASE)
    s = re.sub(r"\s*±\s*", " ± ", s)
    return norm(s)

def is_noise_only_text(s: str, allow_math: bool = False) -> bool:
    s = norm(s)
    if not s:
        return True
    if allow_math and any(g in s for g in MATH_GLYPH_MAP.keys()):
        return False
    if s in NOISE_GLYPHS:
        return True
    if re.search(r"[A-Za-z0-9]", s):
        return False
    if PUA_RE.search(s):
        return True
    if PURE_SYMBOL_RE.match(s) and len(s) <= MAX_NOISE_SYMBOL_LEN:
        return True
    return False

def should_drop_line(text: str) -> bool:
    t = norm(normalize_special_glyphs(text))
    if not t:
        return True
    if is_noise_only_text(t):
        return True
    if BATCH_SIZE_VALUE_RE.match(t):
        return True
    low = t.lower()
    for pat in HEADER_DROP_PATTERNS:
        if pat.search(low):
            return True
    return False

# ============================================================
# CHECKBOX tagging
# ============================================================
CHECKBOX_GLYPHS = {"", "☐", "□"}
SINGLE_CHECKBOX_RE = re.compile(r"(?i)^\s*(yes|no|n\s*/\s*a|na)\s*[☐□]\s*$")
CHECKBOX_TOKEN_RE = re.compile(r"(?i)\b(yes|no|n\s*/\s*a|na)\b\s*[☐□]")

def checkbox_tag_for_line(text: str) -> Optional[str]:
    t = norm(normalize_special_glyphs(text))
    if not t:
        return None
    if not any(g in t for g in CHECKBOX_GLYPHS):
        return None

    m = SINGLE_CHECKBOX_RE.match(t)
    if m:
        w = m.group(1).lower().replace(" ", "")
        if w == "yes":
            return "YES_CHECKBOX"
        if w == "no":
            return "NO_CHECKBOX"
        if w in {"na", "n/a"}:
            return "NA_CHECKBOX"
        return "CHECKBOX_CHOICE"

    found = [m.group(1).lower().replace(" ", "") for m in CHECKBOX_TOKEN_RE.finditer(t)]
    if not found:
        return "CHECKBOX_CHOICE"

    has_yes = any(x == "yes" for x in found)
    has_no = any(x == "no" for x in found)
    has_na = any(x in {"na", "n/a"} for x in found)

    if has_yes and has_no and has_na:
        return "YES_NO_NA_CHECKBOX"
    if has_yes and has_na:
        return "YES_NA_CHECKBOX"
    if has_yes and has_no:
        return "YES_NO_CHECKBOX"
    if has_no and has_na:
        return "NO_NA_CHECKBOX"
    if has_yes:
        return "YES_CHECKBOX"
    if has_no:
        return "NO_CHECKBOX"
    if has_na:
        return "NA_CHECKBOX"
    return "CHECKBOX_CHOICE"

# ============================================================
# Instruction tagging
# ============================================================
VERB_START_SET = {
    "add", "mix", "transfer", "record", "verify", "check", "ensure", "weigh", "sieve", "start", "stop", "set",
    "clean", "assemble", "disassemble", "label", "store", "measure", "charge", "pour", "stir", "dry", "heat",
    "cool", "coat", "blend", "collect", "discard", "inspect", "calibrate", "open", "close", "remove", "place",
    "continue", "increase", "decrease", "maintain", "allow", "confirm", "note", "adjust", "fill", "empty",
    "rinse", "attach", "detach", "tighten", "loosen", "connect", "disconnect", "load", "unload", "turn",
    "press", "select", "enter", "print", "sign", "monitor", "shake", "spray", "apply", "dispense",
    "document", "complete", "initial", "refer", "inform"
}
LEADING_BULLET_NUM_RE = re.compile(r"^\s*(?:[-•*]|(\(?\d+[\).])|[A-Za-z][\).])\s+", re.VERBOSE)

def starts_with_verb_instruction(text: str) -> bool:
    t0 = norm(text)
    if not t0:
        return False
    if is_header_label_only(t0):
        return False

    t = re.sub(r"^[^A-Za-z]+", "", t0)
    t = LEADING_BULLET_NUM_RE.sub("", t)

    low = t.lower()
    if low.startswith("do not ") or low.startswith("don't "):
        return True

    m = re.match(r"^[A-Za-z]+", t)
    if not m:
        return False
    return m.group(0).lower() in VERB_START_SET

# ============================================================
# DATE tagging
# ============================================================
DATE_RE = re.compile(
    r"""(?ix)^\s*(
        \d{1,2}[/-]\d{1,2}[/-]\d{2,4} |
        \d{4}[/-]\d{1,2}[/-]\d{1,2} |
        \d{1,2}\s+[A-Za-z]{3,9}\s+\d{2,4} |
        [A-Za-z]{3,9}\s+\d{1,2},\s*\d{4}
    )\s*$"""
)
def is_date_line(text: str) -> bool:
    return bool(DATE_RE.match(norm(text)))

# ============================================================
# Wrapped-line merging
# ============================================================
END_SENT_RE = re.compile(r"[.;:!?)]\s*$")
STARTS_NEW_ITEM_RE = re.compile(r"^\s*(?:[-•*]|\(?\d+[\).]|[A-Za-z][\).])\s+", re.IGNORECASE)

CONTINUATION_WORDS = (
    "and", "or", "to", "of", "the", "a", "an", "in", "at", "on", "for",
    "with", "without", "from", "into", "by", "as", "per", "before", "after",
    "during", "within", "while", "until", "start", "end", "result", "range"
)
CONTINUATION_START_RE = re.compile(
    r"^(?:" + "|".join(map(re.escape, CONTINUATION_WORDS)) + r")\b",
    re.IGNORECASE,
)
END_DANGLING_RE = re.compile(
    r"""(?ix)
    (?:\b(?:""" + "|".join(map(re.escape, CONTINUATION_WORDS)) + r""")\b\s*$)
    |(?:[/:,-]\s*$)
    """
)

def looks_like_continuation(text: str) -> bool:
    t = norm(text)
    if not t:
        return False
    if re.match(r"^[a-z]", t):
        return True
    if CONTINUATION_START_RE.match(t):
        return True
    if re.match(r"^\d+\b", t):
        return True
    if len(t.split()) <= 2 and not STARTS_NEW_ITEM_RE.match(t):
        return True
    return False

@dataclass(frozen=True)
class Word:
    x0: float
    y0: float
    x1: float
    y1: float
    text: str
    block: int
    line: int

@dataclass
class LineObj:
    page: int
    y: float
    x: float
    block: int
    line: int
    text: str
    words: List[Word]
    page_h: float

def is_table_value_like(text: str) -> bool:
    """True for values that should usually stay in their own table cell/row."""
    t = norm(text)
    if not t:
        return False
    if is_quantity_with_unit(t):
        return True
    if re.search(r"[<>≤≥]\s*\d", t):
        return True
    if re.fullmatch(r"\d+(?:[.,]\d+)?\s*(?:rpm|sec|s|min|mins|h|hrs|hz|bar|hpa|kg|g|mg|l|ml|%|°c|c)\b.*", t, re.IGNORECASE):
        return True
    return False


def should_merge_wrapped(prev: LineObj, cur: LineObj) -> bool:
    """
    Merge only real wrapped text from the same visual cell/paragraph.

    This intentionally avoids broad semantic merging because it can join unrelated
    table rows, e.g. Product temperature + Wetting Cycle, or multiple instructions.
    """
    if prev.page != cur.page:
        return False

    p = norm(prev.text)
    c = norm(cur.text)
    if not p or not c:
        return False

    # Never merge headings, sections, stages, checkbox rows, or new bullet/list rows.
    if STAGE_RE.match(p) or STAGE_RE.match(c):
        return False
    if SECTION_RE.match(p) or SECTION_RE.match(c):
        return False
    if NO_STAGE_SECTION_RE.match(c):
        return False
    if STARTS_NEW_ITEM_RE.match(c):
        return False
    if checkbox_tag_for_line(p) or checkbox_tag_for_line(c):
        return False

    # Never merge a new clear instruction into the previous text.
    # Exception: lowercase continuation lines are handled below.
    if starts_with_verb_instruction(c) and not re.match(r"^[a-z]", c):
        return False

    x_diff = abs(prev.x - cur.x)
    y_gap = cur.y - prev.y
    same_block = prev.block == cur.block
    x_close = x_diff <= 18.0

    if y_gap < -1 or y_gap > 24.0:
        return False

    # Avoid merging separate table value rows/cells such as:
    # Product temperature / at end ≥ 38 °C / The quantity...
    if is_table_value_like(c) and not END_DANGLING_RE.search(p):
        return False

    continuation = (
        re.match(r"^[a-z]", c) is not None
        or CONTINUATION_START_RE.match(c) is not None
        or END_DANGLING_RE.search(p) is not None
    )
    if not continuation:
        return False

    # If previous sentence is finished, only merge when the next line is clearly
    # lowercase continuation. This prevents joining two different instructions.
    if END_SENT_RE.search(p) and not re.match(r"^[a-z]", c):
        return False

    # Main safe case: same PyMuPDF block and same visual x-position.
    if same_block and x_close:
        return True

    # PIP/material tables often split one cell into adjacent PDF blocks with
    # shifted x positions. Allow lowercase continuation across nearby blocks.
    # This merges examples like:
    #   "Verify correct Flavonoid Lot Numbers are" +
    #   "staged in the granulation room."
    # and:
    #   "The quantity of purified water" +
    #   "is dictated by the recipe that was selected ..."
    if re.match(r"^[a-z]", c) and x_diff <= 90.0 and y_gap <= 24.0:
        return True

    # Some PDFs split a single cell across adjacent blocks. Permit when
    # previous line ends dangling, such as "at", "of", "are", comma, hyphen.
    if END_DANGLING_RE.search(p) and x_diff <= 160.0 and y_gap <= 24.0:
        return True

    return False

def merge_wrapped_lines(lines: List[LineObj]) -> List[LineObj]:
    if not lines:
        return lines
    lines = sorted(lines, key=lambda ln: (ln.page, ln.y, ln.x, ln.block, ln.line))
    merged: List[LineObj] = []
    cur = lines[0]
    for nxt in lines[1:]:
        if should_merge_wrapped(cur, nxt):
            cur.text = norm(cur.text + " " + nxt.text)
            cur.words = (cur.words or []) + (nxt.words or [])
        else:
            merged.append(cur)
            cur = nxt
    merged.append(cur)
    return merged

# ============================================================
# Drop standalone row-index numbers
# ============================================================
SEQ_NUM_LINE_RE = re.compile(r"^\s*(\d{1,3})(?:\s+(\d{1,3}))?\s*$")
def is_table_row_index_line(ln: LineObj) -> bool:
    t = norm(ln.text)
    m = SEQ_NUM_LINE_RE.match(t)
    if not m:
        return False
    a = int(m.group(1))
    b = int(m.group(2)) if m.group(2) else None
    if a <= 0 or a > 500:
        return False
    if b is not None and (b <= 0 or b > 500):
        return False
    if ln.x > 120:
        return False
    return True

# ============================================================
# ✅ IPC header detection (span tagging) + structured emission
# ============================================================
IPC_FIELDS = [
    "TEST_NO", "DATE", "IPC_TEST_TIME", "IPC_RESULT",
    "APPEARANCE_TEST_TIME", "DEFECTS", "APPEARANCE_RESULT", "SIGN"
]
IPC_HEADER_LABELS = {
    "TEST_NO": "Test No",
    "DATE": "Date",
    "IPC_TEST_TIME": "IPC Test Time",
    "IPC_RESULT": "IPC Result",
    "APPEARANCE_TEST_TIME": "Appearance Test Time",
    "DEFECTS": "*Defects (Number and type)",
    "APPEARANCE_RESULT": "Appearance Result",
    "SIGN": "Sign",
}

IPC_HEADER_FRAGMENT_RE = re.compile(
    r"""(?ix)
    ^\s*(
        \*?\s*defects? |
        ipc(\s+test)? |
        appearance |
        test\s*no\.? |
        date |
        sign |
        time |
        result |
        test\s*time |
        \(\s*number\s+and |
        number\s+and |
        type\)? |
        \)\s*$
    )\s*$
    """
)

def is_ipc_header_fragment(text: str) -> bool:
    t = norm(text)
    if not t:
        return False
    return bool(IPC_HEADER_FRAGMENT_RE.match(t))

def tokens_lower(s: str) -> Set[str]:
    return set(re.findall(r"[a-z]+", (s or "").lower()))

def _ipc_header_score(tok: Set[str]) -> int:
    s_test = 1 if ("test" in tok and ("no" in tok or "number" in tok)) else 0
    s_ipc = 1 if ("ipc" in tok) else 0
    s_app = 1 if ("appearance" in tok) else 0
    s_def = 1 if ("defects" in tok or "defect" in tok) else 0
    s_res = 1 if ("result" in tok) else 0
    s_time = 1 if ("time" in tok) else 0
    s_sign = 1 if ("sign" in tok or "signature" in tok) else 0
    return s_test + s_ipc + s_app + s_def + s_res + s_time + s_sign

def _is_ipc_test_header_tokens_relaxed(tok: Set[str]) -> bool:
    has_test = ("test" in tok and ("no" in tok or "number" in tok))
    has_ipc = ("ipc" in tok)
    return has_test and has_ipc and (_ipc_header_score(tok) >= 4)

def detect_ipc_test_header(
    lines_on_page: List[LineObj],
    idx: int,
    lookahead: int = 18
) -> Optional[Tuple[int, int]]:
    merged: Set[str] = set()
    header_start: Optional[int] = None
    best_end: Optional[int] = None
    best_score = 0

    for j in range(idx, min(len(lines_on_page), idx + lookahead)):
        line_tok = tokens_lower(lines_on_page[j].text)
        merged |= line_tok

        if header_start is None and ("test" in line_tok) and (("no" in line_tok) or ("number" in line_tok)):
            header_start = j

        sc = _ipc_header_score(merged)
        if sc > best_score:
            best_score = sc
            best_end = j

        if _is_ipc_test_header_tokens_relaxed(merged):
            start = header_start if header_start is not None else idx
            return (start, j)

    if best_score >= 5 and header_start is not None and best_end is not None:
        return (header_start, best_end)
    return None

def build_ipc_header_spans(by_page: Dict[int, List[LineObj]]) -> Dict[int, List[Tuple[int, int]]]:
    spans: Dict[int, List[Tuple[int, int]]] = {}
    for p, lst in by_page.items():
        spans[p] = []
        i = 0
        while i < len(lst):
            ih = detect_ipc_test_header(lst, i)
            if ih:
                spans[p].append(ih)
                i = ih[1] + 1
                continue
            i += 1
    return spans

def is_index_in_any_span(i: int, spans_list: List[Tuple[int, int]]) -> bool:
    for a, b in spans_list:
        if a <= i <= b:
            return True
    return False

# ============================================================
# Extract lines with attached words (single pass)
# Also returns doc_name (from "Document name ...")
# ============================================================
DOC_NAME_RE = re.compile(r"(?i)^\s*document\s*name\s+(.+?)\s*$")

def extract_lines_dict_words(pdf_path: str, debug_skip: bool = False) -> Tuple[List[LineObj], List[int], Optional[str]]:
    doc = fitz.open(pdf_path)
    out: List[LineObj] = []
    skipped_pages: List[int] = []
    doc_name: Optional[str] = None
    try:
        for pidx in range(len(doc)):
            page = doc[pidx]
            page_no = pidx + 1
            page_h = float(page.rect.height)

            page_text = page.get_text("text") or ""
            if page_is_docusign_certificate(page_text):
                skipped_pages.append(page_no)
                if debug_skip:
                    print(f"Skipping DocuSign certificate page: {page_no}")
                continue

            words_raw = page.get_text("words") or []
            d = page.get_text("dict") or {}
            blocks = d.get("blocks", [])

            words_by_bl: Dict[Tuple[int, int], List[Word]] = {}
            for (wx0, wy0, wx1, wy1, wtxt, wblock, wline, _wno) in words_raw:
                words_by_bl.setdefault((int(wblock), int(wline)), []).append(
                    Word(float(wx0), float(wy0), float(wx1), float(wy1), str(wtxt), int(wblock), int(wline))
                )

            for bi, b in enumerate(blocks):
                if b.get("type") != 0:
                    continue
                for li, line in enumerate(b.get("lines", []) or []):
                    spans = line.get("spans", [])
                    txt_raw = " ".join(s.get("text", "") for s in spans)
                    txt = normalize_special_glyphs(txt_raw)

                    if doc_name is None:
                        m = DOC_NAME_RE.match(norm(txt))
                        if m:
                            doc_name = norm(m.group(1))

                    if should_drop_line(txt):
                        continue

                    x0, y0, x1, y1 = line.get("bbox", [0, 0, 0, 0])
                    ws = words_by_bl.get((bi, li), [])
                    out.append(
                        LineObj(
                            page=page_no,
                            y=float(y0),
                            x=float(x0),
                            block=int(bi),
                            line=int(li),
                            text=txt,
                            words=ws,
                            page_h=page_h,
                        )
                    )
    finally:
        doc.close()

    return out, skipped_pages, doc_name

# ============================================================
# Repeating header/footer removal (conservative)
# ============================================================
CONTENT_KEYWORDS_RE = re.compile(
    r"\b("
    r"in-?process|control|checks?|station|perform|physical|uniformity|pharmacopoeia|"
    r"ipc|appearance|defects?|test\s*no|"
    r"performed\s*by|prepared\s*by|reviewed\s*by|approved\s*by|author|"
    r"sign|signature|date|"
    r"complete\s+the\s+box"
    r")\b",
    re.IGNORECASE,
)

def is_probably_header_or_footer(ln: LineObj, band: float = 80.0) -> bool:
    return (ln.y <= band) or (ln.y >= (ln.page_h - band))

def find_repeating_lines_conservative(lines: List[LineObj], min_pages: int) -> Set[str]:
    pages_by_text: Dict[str, Set[int]] = {}
    for ln in lines:
        t = ln.text.strip()
        if not t:
            continue
        k = t.lower()

        if not (is_probably_header_or_footer(ln) or any(p.search(k) for p in HEADER_DROP_PATTERNS)):
            continue
        if CONTENT_KEYWORDS_RE.search(t):
            continue

        pages_by_text.setdefault(k, set()).add(ln.page)

    return {k for k, ps in pages_by_text.items() if len(ps) >= min_pages}

def drop_batchsize_header_block(lines: List[LineObj]) -> List[LineObj]:
    out: List[LineObj] = []
    i = 0
    n = len(lines)
    while i < n:
        if lines[i].text.strip().lower() == "new document":
            page = lines[i].page
            found_bs = None
            found_val = None
            for j in range(i + 1, min(n, i + 7)):
                if lines[j].page != page:
                    break
                tj = lines[j].text.strip()
                if tj.lower() == "batch size":
                    found_bs = j
                if BATCH_SIZE_VALUE_RE.match(tj):
                    found_val = j
            if found_bs is not None and found_val is not None:
                i = max(found_bs, found_val) + 1
                continue
        out.append(lines[i])
        i += 1
    return out

# ============================================================
# TABLE template detection (Materials + PIP + Routing)
# ============================================================
VPR_RE = re.compile(r"validated\s+product\s+routing", re.IGNORECASE)
EQUIP_TEMPLATE_RE = re.compile(r"equipment\s+template", re.IGNORECASE)

EQUIP_CODE_RE = re.compile(r"\b[A-Z]{2}\.[A-Z]\.[A-Z]{3}\.[A-Z0-9]{4,}\b")
PPMMAN_TEMPLATE_RE = re.compile(r"\bPP\.M\.MAN\.[A-Z0-9]{3,}(?:\.[A-Z0-9]+)*\b", re.IGNORECASE)

def is_equipment_template_code(s: str) -> bool:
    t = norm(s)
    return bool(EQUIP_CODE_RE.search(t) or PPMMAN_TEMPLATE_RE.search(t))

def tokens(s: str) -> Set[str]:
    return set(re.findall(r"[a-z]+", (s or "").lower()))

def _is_material_header_tokens(tok: Set[str]) -> bool:
    has_mat = ("material" in tok or "materials" in tok) and ("name" in tok)
    has_item = ("item" in tok) and ("code" in tok)
    has_prep = ("prep" in tok or "preparation" in tok) and ("code" in tok)
    has_qty = ("quantity" in tok) and ("required" in tok or "require" in tok)
    return has_mat and has_item and has_prep and has_qty

def _is_pip_header_tokens(tok: Set[str]) -> bool:
    has_pi = ("process" in tok) and (("instruction" in tok) or ("instructions" in tok))
    has_req = ("requirements" in tok) or ("requirement" in tok)
    has_param = ("parameters" in tok) or ("parameter" in tok)
    return has_pi and has_req and has_param

def detect_material_header(lines_on_page: List[LineObj], idx: int, lookahead: int = 8) -> Optional[Tuple[int, int]]:
    merged: Set[str] = set()
    for j in range(idx, min(len(lines_on_page), idx + lookahead)):
        merged |= tokens(lines_on_page[j].text)
        if _is_material_header_tokens(merged):
            return (idx, j)
    return None

def detect_pip_header(lines_on_page: List[LineObj], idx: int, lookahead: int = 6) -> Optional[Tuple[int, int]]:
    merged: Set[str] = set()
    for j in range(idx, min(len(lines_on_page), idx + lookahead)):
        merged |= tokens(lines_on_page[j].text)
        if _is_pip_header_tokens(merged):
            return (idx, j)
    return None

def infer_boundaries_from_words(header_words: List[Word]) -> List[float]:
    ws = sorted(header_words, key=lambda w: w.x0)
    if len(ws) <= 1:
        return []
    gaps = []
    for a, b in zip(ws, ws[1:]):
        gaps.append(b.x0 - a.x1)
    pos_gaps = [g for g in gaps if g > 0]
    if not pos_gaps:
        return []
    pos_gaps_sorted = sorted(pos_gaps)
    med = pos_gaps_sorted[len(pos_gaps_sorted) // 2]
    thr = max(12.0, med * 2.5)

    boundaries = []
    for (a, b) in zip(ws, ws[1:]):
        g = b.x0 - a.x1
        if g >= thr:
            boundaries.append((a.x1 + b.x0) / 2.0)
    return sorted(boundaries)

def split_cells(words: List[Word], boundaries: List[float]) -> List[str]:
    if not words:
        return []

    # Assign to columns by x, but preserve natural reading order inside each cell.
    # The old version sorted only by x, which changed wrapped text like:
    #   "... at start of the" + "Granulation."
    # into "Dispense Granulation. Flavonoide ...".
    cols = [[] for _ in range(len(boundaries) + 1)]
    for w in sorted(words, key=lambda z: (z.y0, z.x0)):
        k = 0
        while k < len(boundaries) and w.x0 > boundaries[k]:
            k += 1
        cols[k].append(w)

    return [norm(" ".join(w.text for w in sorted(col, key=lambda z: (z.y0, z.x0)))) for col in cols]

@dataclass
class Schema:
    kind: str
    boundaries: List[float]
    col_tags: List[str]

def build_page_schemas(lines: List[LineObj]) -> Dict[int, List[Tuple[float, Schema]]]:
    by_page: Dict[int, List[LineObj]] = {}
    for ln in lines:
        by_page.setdefault(ln.page, []).append(ln)

    page_schemas: Dict[int, List[Tuple[float, Schema]]] = {}

    for p, lst in by_page.items():
        lst = sorted(lst, key=lambda x: (x.y, x.x))
        schemas_for_page: List[Tuple[float, Schema]] = []

        i = 0
        while i < len(lst):
            mh = detect_material_header(lst, i)
            if mh:
                start_i, end_i = mh
                header_words: List[Word] = []
                for k in range(start_i, end_i + 1):
                    header_words.extend(lst[k].words or [])
                bounds = infer_boundaries_from_words(header_words)
                schemas_for_page.append(
                    (lst[start_i].y, Schema(kind="MATERIALS", boundaries=bounds,
                                            col_tags=["MATERIAL_NAME", "ITEM_CODE", "PREP_CODE", "QUANTITY_REQUIRED"]))
                )
                i = end_i + 1
                continue

            ph = detect_pip_header(lst, i)
            if ph:
                start_i, end_i = ph
                header_words = []
                for k in range(start_i, end_i + 1):
                    header_words.extend(lst[k].words or [])
                bounds = infer_boundaries_from_words(header_words)
                schemas_for_page.append(
                    (lst[start_i].y, Schema(kind="PIP", boundaries=bounds,
                                            col_tags=["PROCESS_INSTRUCTION", "REQUIREMENTS", "TEXT"]))
                )
                i = end_i + 1
                continue

            # ROUTING presence hint
            if VPR_RE.search(lst[i].text.strip()) or EQUIP_TEMPLATE_RE.search(lst[i].text.strip()):
                schemas_for_page.append((lst[i].y, Schema(kind="ROUTING", boundaries=[], col_tags=[])))

            i += 1

        schemas_for_page = sorted(schemas_for_page, key=lambda t: t[0])
        cleaned: List[Tuple[float, Schema]] = []
        last_y = None
        for y, sc in schemas_for_page:
            if last_y is None or abs(y - last_y) > 2.0:
                cleaned.append((y, sc))
                last_y = y
        page_schemas[p] = cleaned

    return page_schemas

def schema_for_line(page_schemas: Dict[int, List[Tuple[float, Schema]]], ln: LineObj) -> Optional[Schema]:
    sch_list = page_schemas.get(ln.page, [])
    active = None
    for y0, sc in sch_list:
        if ln.y >= y0:
            active = sc
        else:
            break
    return active

# ============================================================
# QUANTITY REQUIRED tagging
# ============================================================
QUANTITY_WITH_UNIT_RE = re.compile(
    r"""(?ix)^\s*
    [-+]?
    (?:\d{1,3}(?:,\d{3})+|\d+)
    (?:\.\d+)?
    \s*
    (?:
        kg|g|mg|mcg|µg|ug|
        l|ml|mL|L|
        %|
        tablets?|capsules?|
        mins?|min|h|hr|hrs|
        °c|c
    )
    \b
    .*
    $
    """
)

def is_quantity_with_unit(text: str) -> bool:
    return bool(QUANTITY_WITH_UNIT_RE.match(norm(text)))

ITEM_CODE_RE2 = re.compile(r"^(?=.*[A-Za-z])(?=.*\d)[A-Za-z0-9]{4,}$")
PREP_CODE_RE2 = re.compile(r"^\d+$")

def looks_like_item_code(s: str) -> bool:
    s = norm(s)
    if PREP_CODE_RE2.match(s):
        return False
    return bool(ITEM_CODE_RE2.match(s))

def looks_like_prep_code(s: str) -> bool:
    return bool(PREP_CODE_RE2.match(norm(s)))

def retag_materials_row_by_content(cells: List[str], col_tags: List[str]) -> List[str]:
    cells = [norm(c) for c in cells]
    if not any(cells):
        return ["TEXT"] * len(cells)

    tags: List[str] = []
    for i, c in enumerate(cells):
        if not c:
            tags.append("TEXT")
            continue

        if is_quantity_with_unit(c):
            tags.append("QUANTITY_REQUIRED")
            continue

        if i == 0 and starts_with_verb_instruction(c):
            tags.append("INSTRUCTION")
            continue

        if looks_like_item_code(c):
            tags.append("ITEM_CODE")
            continue
        if looks_like_prep_code(c):
            tags.append("PREP_CODE")
            continue

        base = col_tags[i] if i < len(col_tags) else "TEXT"
        tags.append(base if base in ALLOWED_TAG_DATA else "TEXT")

    return tags

# ============================================================
# Formula blocks
# ============================================================
CALC_ANCHORS = re.compile(
    r"(?i)\b("
    r"compression calculations|"
    r"tablets produced|"
    r"yield\s*=|"
    r"use the formula|"
    r"enter calculation below|"
    r"net\s*weight|"
    r"average\s*tablet\s*weight|"
    r"1,000,000"
    r")\b"
)
CALC_STOP = re.compile(r"(?i)\b(observations|comments|checklist|reviewed by|approved by)\b")

def is_calc_anchor(text: str) -> bool:
    return bool(CALC_ANCHORS.search(normalize_math_glyphs(text)))

DENOM_RE = re.compile(r"^\s*1\s*,\s*0{6}\s*$")
X_FIX_RE = re.compile(r"[×𝐱𝑥x]", re.IGNORECASE)

FORMULA_TEXT_CONT_RE = re.compile(
    r"""(?ix)\b(
        net\s*weight|
        of\s*tablets?|
        tablets?\s*produced|
        quantity\s*of\s*conforming|
        average\s*tablet\s*weight|
        tablet\s*weight|
        mean\s*weight|
        press\s*speed|
        production\s*since\s*last\s*good\s*test|
        refer\s*sop|
        use\s*the\s*formula|
        enter\s*calculation|
        reconciliation
    )\b"""
)
UNITS_ONLY_RE = re.compile(r"(?ix)^\s*(kg|mg|g|min|mins|tablets|%)\s*\)?\s*$")

def _has_math_signal(t: str) -> bool:
    return bool(re.search(r"(?i)(=|/|\*|\bx\s*100\b|%|\bkg\b|\bmg\b|\bmin\b|\b\d+(?:[.,]\d+)?\b)", t))

def looks_like_calc_continuation(text: str) -> bool:
    t = normalize_math_glyphs(text)
    if _has_math_signal(t):
        return True
    if FORMULA_TEXT_CONT_RE.search(t):
        return True
    if UNITS_ONLY_RE.match(t):
        return True
    if re.match(r"(?ix)^\s*\(\s*[a-z/ ]+\s*\)\s*$", t):
        return True
    return False

def normalize_formula_block(calc_lines: List[str]) -> str:
    lines = [normalize_special_glyphs(normalize_math_glyphs(x)) for x in (calc_lines or []) if norm(x)]
    if not lines:
        return ""
    lines = [X_FIX_RE.sub(" x ", ln) for ln in lines]
    lines = [norm(ln) for ln in lines]

    if len(lines) >= 2 and DENOM_RE.match(lines[-1]):
        denom = "1,000,000"
        expr = " ".join(lines[:-1])
        if re.search(r"/\s*1\s*,\s*0{6}\b", expr):
            return norm(expr)
        return norm(f"{expr} / {denom}")

    return norm(" ".join(lines))

# ============================================================
# Needs-review columns
# ============================================================
BULLET_RE = re.compile(r"^\s*[-•*]\s+")
LINEBREAK_RE = re.compile(r"[\r\n]+")
MANY_SPACES_RE = re.compile(r"\s{6,}")

def _safe_str(x) -> str:
    return "" if x is None else str(x)

def add_needs_review_columns(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()

    title = out["Data title"].map(_safe_str) if "Data title" in out.columns else pd.Series("", index=out.index)
    proc = out["Process step"].map(_safe_str) if "Process step" in out.columns else pd.Series("", index=out.index)
    subp = out["Sub process step"].map(_safe_str) if "Sub process step" in out.columns else pd.Series("", index=out.index)
    page = out["Page number"] if "Page number" in out.columns else pd.Series(None, index=out.index)
    para = out["Paragraph number"] if "Paragraph number" in out.columns else pd.Series(None, index=out.index)
    tagd = out["Tag data"].map(_safe_str).str.upper() if "Tag data" in out.columns else pd.Series("", index=out.index)

    needs = pd.Series(False, index=out.index)
    reasons = pd.Series("", index=out.index)

    def flag(mask: pd.Series, reason: str):
        nonlocal needs, reasons
        mask = mask.fillna(False)
        if mask.any():
            needs = needs | mask
            reasons = reasons.where(~mask, (reasons + f"|{reason}").str.strip("|"))

    pnum = pd.to_numeric(page, errors="coerce")
    parnum = pd.to_numeric(para, errors="coerce")

    flag(pnum.isna() | (pnum < 1) | (pnum % 1 != 0), "invalid_page_number")
    flag(parnum.isna() | (parnum < 0) | (parnum % 1 != 0), "invalid_paragraph_number")

    if set(["Page number", "Paragraph number", "Data title", "Tag data"]).issubset(out.columns):
        dup_mask = out.duplicated(subset=["Page number", "Paragraph number", "Data title", "Tag data"], keep=False)
        flag(dup_mask, "possible_duplicate")

    flag(title.str.strip().eq(""), "missing_data_title")
    flag(title.str.contains(LINEBREAK_RE), "multiline_text")
    flag(title.str.match(BULLET_RE), "bullet_list")
    flag(title.str.len() > 200, "very_long_text")
    flag(title.str.contains(MANY_SPACES_RE), "layout_noise")

    flag(proc.str.strip().eq("UNKNOWN_PROCESS"), "unknown_process")
    flag((proc.str.strip().ne("")) & (proc.str.strip().ne("--")) & (subp.str.strip().eq("")), "subprocess_missing")

    flag(tagd.eq("FORMULA"), "formula_needs_review")

    instr_text = title.fillna("").map(_safe_str).str.strip()
    instr_is_real = instr_text.map(starts_with_verb_instruction)
    instr_no_period = instr_is_real & instr_text.ne("") & ~instr_text.str.endswith(".")
    flag(instr_no_period, "instruction_not_ended_with_dot")

    out["Needs review"] = needs.map(lambda x: "Yes" if x else "No")
    out["Review reasons"] = reasons.replace("", "-")
    return out

# ============================================================
# Label splitting + signature fixes
# ============================================================
MERGE_LABELS = ["Reviewed By", "Approved By", "Prepared By"]
MERGE_LABEL_RE = re.compile(r"(?i)\b(" + "|".join(re.escape(x) for x in MERGE_LABELS) + r")\b")

def split_merged_label_runs(text: str) -> List[str]:
    t = norm(text)
    if not t:
        return []

    matches = list(MERGE_LABEL_RE.finditer(t))
    if not matches:
        return [t]

    if len(matches) == 1 and matches[0].start() == 0:
        return [t]

    spans = [(m.start(), m.end(), m.group(1)) for m in matches]
    spans.append((len(t), len(t), None))

    prefix = t[:spans[0][0]].strip()
    segments: List[str] = []

    for i in range(len(spans) - 1):
        start, end, label = spans[i]
        next_start = spans[i + 1][0]
        after = t[end:next_start].strip()

        if not label:
            continue

        lab = norm(label)

        if i == 0 and prefix:
            segments.append(norm(prefix))
            prefix = ""

        if after:
            segments.append(norm(f"{lab} {after}"))
        else:
            segments.append(lab)

    segments = [s for s in segments if s]
    return segments if segments else [t]

LABELS_CAN_SPLIT = {"Reviewed By", "Approved By"}

def split_label_and_role_if_combined(text: str) -> List[str]:
    t = norm(text)
    if not t:
        return []
    low = t.lower()
    for lab in LABELS_CAN_SPLIT:
        l = lab.lower()
        if low.startswith(l + " "):
            role = norm(t[len(lab):]).lstrip()
            if role:
                return [lab, role]
            return [lab]
    return [t]

def is_exact_label(text: str, label: str) -> bool:
    return norm(text).strip().lower() == label.strip().lower()

def starts_with_label(text: str, label: str) -> bool:
    t = norm(text).strip().lower()
    l = label.strip().lower()
    return t == l or t.startswith(l + " ")

def split_trailing_title(text: str) -> List[str]:
    t = norm(text)
    if not t:
        return []
    if t.lower().endswith(" title") and t.lower() != "title":
        left = norm(t[:-len(" title")])
        if left:
            return [left, "Title"]
    return [t]

def is_plain_sig_word(s: str, word: str) -> bool:
    t = norm(s).lower()
    return t == word.lower()

# ============================================================
# ✅ Paragraph anchors (auto page detection)
# ============================================================
RX_VPR_ANCHOR = re.compile(r"validated\s+product\s+routing", re.IGNORECASE)
RX_PREPARED_BY = re.compile(r"\bprepared\s+by\b", re.IGNORECASE)
RX_CONFIRM = re.compile(r"^\s*confirm\b", re.IGNORECASE)
RX_RAW_MAT = re.compile(r"raw\s+materials\s+required", re.IGNORECASE)

def _find_first_page_with(lines: List[LineObj], rx: re.Pattern) -> Optional[int]:
    for ln in sorted(lines, key=lambda z: (z.page, z.y, z.x)):
        if rx.search(norm(ln.text)):
            return ln.page
    return None

def _find_anchor_y_on_page(page_lines: List[LineObj], rx: re.Pattern) -> Optional[float]:
    for ln in sorted(page_lines, key=lambda z: (z.y, z.x)):
        if rx.search(norm(ln.text)):
            return float(ln.y)
    return None

def compute_paragraph_number(
    page: int,
    y: float,
    routing_page: Optional[int],
    y_prepared: Optional[float],
    y_confirm: Optional[float],
    materials_page: Optional[int],
    y_raw: Optional[float],
) -> int:
    # Routing page: 3 blocks
    if routing_page is not None and page == routing_page and y_prepared is not None and y_confirm is not None:
        if y < y_prepared:
            return 1
        if y < y_confirm:
            return 2
        return 3

    # Materials page: 2 blocks
    if materials_page is not None and page == materials_page and y_raw is not None:
        return 1 if y < y_raw else 2

    return 1

# ============================================================
# MAIN export
# ============================================================
def extract_to_excel(
    pdf_path: str,
    out_xlsx: str,
    debug_skip_pages: bool = False,
    product_name: Optional[str] = None,
) -> pd.DataFrame:
    pdf_name = os.path.basename(pdf_path)

    lines, skipped_pages, doc_name = extract_lines_dict_words(pdf_path, debug_skip=debug_skip_pages)
    lines = merge_wrapped_lines(lines)
    lines = [l for l in lines if not should_drop_line(l.text)]

    has_any_stage = any(STAGE_RE.match((l.text or "").strip()) for l in lines)

    min_pages = 2 if len({l.page for l in lines}) <= 6 else 3
    repeating = find_repeating_lines_conservative(lines, min_pages=min_pages)
    lines = [l for l in lines if l.text.strip().lower() not in repeating]
    lines = drop_batchsize_header_block(lines)

    process_step = norm(product_name) if product_name and norm(product_name) else "--"

    by_page: Dict[int, List[LineObj]] = {}
    for ln in lines:
        by_page.setdefault(ln.page, []).append(ln)
    for p in by_page:
        by_page[p] = sorted(by_page[p], key=lambda x: (x.y, x.x))

    page_schemas = build_page_schemas(lines)

    # ============================================================
    # ✅ IPC header span detection per page
    # ============================================================
    ipc_header_spans = build_ipc_header_spans(by_page)
    ipc_pages: Set[int] = {p for p, spans in ipc_header_spans.items() if spans}

    idx_map: Dict[int, Dict[int, int]] = {
        p: {id(obj): i for i, obj in enumerate(lst)} for p, lst in by_page.items()
    }

    ipc_header_idx_flags: Dict[int, List[bool]] = {}
    for p, lst in by_page.items():
        spans_list = ipc_header_spans.get(p, [])
        flags = [False] * len(lst)
        if spans_list:
            for i in range(len(lst)):
                if is_index_in_any_span(i, spans_list):
                    flags[i] = True
        ipc_header_idx_flags[p] = flags

    emitted_ipc_header_pages: Set[int] = set()

    # ============================================================
    # ✅ Auto-detect pages + Y anchors for paragraph splits
    # ============================================================
    routing_page = _find_first_page_with(lines, RX_VPR_ANCHOR)
    y_prepared = None
    y_confirm = None
    if routing_page is not None:
        y_prepared = _find_anchor_y_on_page(by_page.get(routing_page, []), RX_PREPARED_BY)
        y_confirm = _find_anchor_y_on_page(by_page.get(routing_page, []), RX_CONFIRM)

        # if "Confirm" appears above "Prepared By" due to extraction oddities, fix order safely
        if y_prepared is not None and y_confirm is not None and y_confirm < y_prepared:
            below = [
                ln for ln in by_page.get(routing_page, [])
                if float(ln.y) > y_prepared and RX_CONFIRM.search(norm(ln.text))
            ]
            if below:
                y_confirm = float(sorted(below, key=lambda z: z.y)[0].y)

    materials_page = _find_first_page_with(lines, RX_RAW_MAT)
    y_raw = None
    if materials_page is not None:
        y_raw = _find_anchor_y_on_page(by_page.get(materials_page, []), RX_RAW_MAT)

    # ============================================================
    # Extraction state
    # ============================================================
    rows: List[Dict] = []
    data_number = 0

    current_sub = ""
    if (not has_any_stage) and doc_name:
        current_sub = doc_name

    routing_state: Dict[int, Dict[str, Optional[str]]] = {}

    calc_active = False
    calc_page: Optional[int] = None
    calc_lines: List[str] = []
    last_calc_y = None

    def flush_calc():
        nonlocal data_number, calc_active, calc_page, calc_lines, last_calc_y
        if calc_active and calc_lines and calc_page is not None:
            data_number += 1
            rows.append({
                "PDF name": pdf_name,
                "Data number": data_number,
                "Page number": calc_page,
                "Paragraph number": 1,  # overridden by compute_paragraph_number when output is built
                "Process step": process_step,
                "Sub process step": current_sub if current_sub else "--",
                "Data title": normalize_formula_block(calc_lines),
                "Tag data": "FORMULA",
            })
        calc_active = False
        calc_page = None
        calc_lines = []
        last_calc_y = None

    def process_one_line(ln: LineObj):
        nonlocal data_number, current_sub, calc_active, calc_page, calc_lines, last_calc_y

        txt0 = (ln.text or "").strip()

        # SUB PROCESS STEP LOGIC
        m = STAGE_RE.match(txt0)
        if m:
            stage_name = norm_sub_process_step(m.group(2))
            if stage_name:
                current_sub = stage_name

        if (not has_any_stage):
            if NO_STAGE_SECTION_RE.match(txt0):
                current_sub = norm(txt0)
            elif CAPS_SECTION_RE.match(txt0):
                up = norm(txt0).upper()
                if not any(b in up for b in CAPS_SECTION_BLOCKLIST):
                    current_sub = norm(txt0)

        # FORMULA detection
        t_math = normalize_math_glyphs(ln.text)
        if CALC_STOP.search(t_math):
            if calc_active and calc_page == ln.page:
                flush_calc()

        if is_calc_anchor(t_math) or (calc_active and calc_page == ln.page and looks_like_calc_continuation(t_math)):
            if (not calc_active) or (calc_page != ln.page):
                flush_calc()
                calc_active = True
                calc_page = ln.page
                calc_lines = [t_math]
                last_calc_y = ln.y
            else:
                y_gap = (ln.y - (last_calc_y if last_calc_y is not None else ln.y))
                if y_gap <= 40 or looks_like_calc_continuation(t_math):
                    calc_lines.append(t_math)
                    last_calc_y = ln.y
                else:
                    flush_calc()
                    calc_active = True
                    calc_page = ln.page
                    calc_lines = [t_math]
                    last_calc_y = ln.y
            return
        else:
            if calc_active and calc_page == ln.page:
                flush_calc()

        if is_table_row_index_line(ln):
            return

        active_schema = schema_for_line(page_schemas, ln)

        # ✅ Paragraph number from anchors (NOT from table logic)
        para_num = compute_paragraph_number(
            page=ln.page,
            y=float(ln.y),
            routing_page=routing_page,
            y_prepared=y_prepared,
            y_confirm=y_confirm,
            materials_page=materials_page,
            y_raw=y_raw,
        )

        # ============================================================
        # ✅ IPC header suppression + structured emission (once per page)
        # ============================================================
        lid = id(ln)
        idx = idx_map.get(ln.page, {}).get(lid)
        line_in_ipc_span = False
        if idx is not None:
            line_in_ipc_span = ipc_header_idx_flags.get(ln.page, [False])[idx]

        if ln.page in ipc_pages:
            trigger = line_in_ipc_span or is_ipc_header_fragment(ln.text)

            if trigger and ln.page not in emitted_ipc_header_pages:
                for field in IPC_FIELDS:
                    data_number += 1
                    rows.append({
                        "PDF name": pdf_name,
                        "Data number": data_number,
                        "Page number": ln.page,
                        "Paragraph number": para_num,
                        "Process step": process_step,
                        "Sub process step": current_sub if current_sub else "--",
                        "Data title": IPC_HEADER_LABELS[field],
                        "Tag data": "IPC_HEADER_STRUCT",
                    })
                emitted_ipc_header_pages.add(ln.page)

            if trigger:
                return

        # ---------------- ROUTING-like tagging (works even if schema misses) ----------------
        routing_like = (routing_page is not None and ln.page == routing_page)

        if (active_schema and active_schema.kind == "ROUTING") or routing_like:
            st = routing_state.setdefault(ln.page, {"mode": None})
            t = ln.text.strip()

            cb_tag = checkbox_tag_for_line(t)
            if cb_tag:
                tag = cb_tag
            else:
                tag = "TEXT"
                if VPR_RE.search(t) or EQUIP_TEMPLATE_RE.search(t):
                    tag = "TEXT"
                elif STAGE_RE.match(t):
                    tag = "VPR_STAGE_NAME"
                    st["mode"] = "EXPECT_EQUIP"
                elif st.get("mode") == "EXPECT_EQUIP":
                    tag = "VPR_EQUIPMENT"
                    st["mode"] = "EXPECT_TEMPLATE"
                elif st.get("mode") == "EXPECT_TEMPLATE":
                    tag = "VPR_EQUIPMENT_TEMPLATE" if is_equipment_template_code(t) else "TEXT"
                    st["mode"] = None
                elif starts_with_verb_instruction(t):
                    tag = "INSTRUCTION"
                elif is_date_line(t):
                    tag = "DATE"
                elif is_quantity_with_unit(t):
                    tag = "QUANTITY_REQUIRED"

            data_number += 1
            rows.append({
                "PDF name": pdf_name,
                "Data number": data_number,
                "Page number": ln.page,
                "Paragraph number": para_num,
                "Process step": process_step,
                "Sub process step": current_sub if current_sub else "--",
                "Data title": ln.text,
                "Tag data": tag_data_safe(tag),
            })
            return

        # ---------------- TABLE block ----------------
        in_table = bool(active_schema and active_schema.kind in {"MATERIALS", "PIP"})
        if in_table:
            # PIP tables are visually 3-column, but PyMuPDF often splits one logical
            # cell/row into several artificial cells. Emit the full reconstructed line
            # so rows like "Product temperature at end ≥ 38 °C" and wrapped
            # requirement text stay together.
            if active_schema.kind == "PIP":
                cb_tag = checkbox_tag_for_line(ln.text)
                if cb_tag:
                    tag = cb_tag
                elif starts_with_verb_instruction(ln.text):
                    tag = "INSTRUCTION"
                elif is_date_line(ln.text):
                    tag = "DATE"
                elif is_quantity_with_unit(ln.text):
                    tag = "QUANTITY_REQUIRED"
                else:
                    tag = "TEXT"

                data_number += 1
                rows.append({
                    "PDF name": pdf_name,
                    "Data number": data_number,
                    "Page number": ln.page,
                    "Paragraph number": para_num,
                    "Process step": process_step,
                    "Sub process step": current_sub if current_sub else "--",
                    "Data title": normalize_special_glyphs(ln.text),
                    "Tag data": tag_data_safe(tag),
                })
                return

            # Full-width instruction/paragraph cells should not be split by table
            # column boundaries. Otherwise a wrapped instruction that crosses a
            # boundary becomes two extracted rows, e.g. "... at" / "start of the".
            if starts_with_verb_instruction(ln.text):
                data_number += 1
                rows.append({
                    "PDF name": pdf_name,
                    "Data number": data_number,
                    "Page number": ln.page,
                    "Paragraph number": para_num,
                    "Process step": process_step,
                    "Sub process step": current_sub if current_sub else "--",
                    "Data title": normalize_special_glyphs(ln.text),
                    "Tag data": "INSTRUCTION",
                })
                return

            cells = split_cells(ln.words or [], active_schema.boundaries)

            if active_schema.kind == "MATERIALS":
                cell_tags = retag_materials_row_by_content(cells, active_schema.col_tags)
            else:
                cell_tags = [
                    active_schema.col_tags[i] if i < len(active_schema.col_tags) else "TEXT"
                    for i in range(len(cells))
                ]

            for c, cell_tag in zip(cells, cell_tags):
                c = normalize_special_glyphs(c)
                if not c:
                    continue

                cb_tag = checkbox_tag_for_line(c)
                if cb_tag:
                    tag = cb_tag
                elif starts_with_verb_instruction(c):
                    tag = "INSTRUCTION"
                elif is_date_line(c):
                    tag = "DATE"
                else:
                    tag = cell_tag

                data_number += 1
                rows.append({
                    "PDF name": pdf_name,
                    "Data number": data_number,
                    "Page number": ln.page,
                    "Paragraph number": para_num,
                    "Process step": process_step,
                    "Sub process step": current_sub if current_sub else "--",
                    "Data title": c,
                    "Tag data": tag_data_safe(tag),
                })
            return

        # ---------------- NORMAL block ----------------
        if is_noise_only_text(ln.text, allow_math=False):
            return

        t = ln.text
        cb_tag = checkbox_tag_for_line(t)
        if cb_tag:
            tag = cb_tag
        elif starts_with_verb_instruction(t):
            tag = "INSTRUCTION"
        elif is_date_line(t):
            tag = "DATE"
        else:
            tag = "TEXT"

        data_number += 1
        rows.append({
            "PDF name": pdf_name,
            "Data number": data_number,
            "Page number": ln.page,
            "Paragraph number": para_num,
            "Process step": process_step,
            "Sub process step": current_sub if current_sub else "--",
            "Data title": ln.text,
            "Tag data": tag_data_safe(tag),
        })

    # ============================================================
    # Main iteration with label splitting fixes
    # ============================================================
    sorted_lines = sorted(lines, key=lambda x: (x.page, x.y, x.x))
    i = 0
    while i < len(sorted_lines):
        ln = sorted_lines[i]

        next_text = ""
        if (i + 1) < len(sorted_lines) and sorted_lines[i + 1].page == ln.page:
            next_text = sorted_lines[i + 1].text

        cur_text = ln.text
        if is_plain_sig_word(cur_text, "reviewed"):
            cur_text = "Reviewed By"
        elif is_plain_sig_word(cur_text, "approved"):
            cur_text = "Approved By"

        if is_exact_label(cur_text, "Reviewed By") and starts_with_label(next_text, "Reviewed By") and not is_exact_label(next_text, "Reviewed By"):
            i += 1
            continue

        expanded: List[str] = []
        for t1 in split_trailing_title(cur_text):
            for t2 in split_merged_label_runs(t1):
                expanded.extend(split_label_and_role_if_combined(t2))

        for out_txt in expanded:
            ln_eff = LineObj(
                page=ln.page, y=ln.y, x=ln.x,
                block=ln.block, line=ln.line,
                text=out_txt, words=ln.words, page_h=ln.page_h
            )
            process_one_line(ln_eff)

        i += 1

    flush_calc()

    for p in skipped_pages:
        rows.append({
            "PDF name": pdf_name,
            "Data number": None,
            "Page number": p,
            "Paragraph number": 0,
            "Process step": "--",
            "Sub process step": "--",
            "Data title": "DOCUSIGN_CERTIFICATE_PAGE_SKIPPED",
            "Tag data": "TEXT",
        })

    # Sort and renumber Data number
    rows.sort(key=lambda r: (
        int(r["Page number"]) if pd.notna(r.get("Page number")) else 10**9,
        int(r["Paragraph number"]) if pd.notna(r.get("Paragraph number")) else 10**9,
    ))
    for j, r in enumerate(rows, start=1):
        r["Data number"] = j

    df = pd.DataFrame(rows, columns=[
        "PDF name",
        "Data number", "Page number", "Paragraph number",
        "Process step", "Sub process step",
        "Data title", "Tag data"
    ])

    df = add_needs_review_columns(df)
    df.to_excel(out_xlsx, index=False, sheet_name="Sheet1")
    return df

# ============================================================
# UI extraction (Streamlit-compatible)
# - PDF name MUST be input PDF name (first row only)
# - Tag column is "Tag" (renamed from Tag data)
# ============================================================
def extract(file_bytes: bytes, file_name: str, product_name: Optional[str] = None) -> pd.DataFrame:
    ext = os.path.splitext(file_name or "")[1].lower()
    if ext != ".pdf":
        raise ValueError("Extractor expects a PDF file.")

    with tempfile.TemporaryDirectory() as td:
        pdf_path = os.path.join(td, "input.pdf")
        with open(pdf_path, "wb") as f:
            f.write(file_bytes)

        tmp_xlsx = os.path.join(td, "out.xlsx")
        df = extract_to_excel(
            pdf_path=pdf_path,
            out_xlsx=tmp_xlsx,
            debug_skip_pages=False,
            product_name=product_name,
        )

        # Build UI dataframe with required columns
        ui_df = df.rename(columns={"Tag data": "Tag"}).copy()
        ui_df.columns = [str(c).strip() for c in ui_df.columns]

        # ✅ guarantee PDF name exists
        if "PDF name" not in ui_df.columns:
            ui_df["PDF name"] = None

        # guarantee other UI columns exist
        for col in ["Process step", "Sub process step", "Tag", "Data title", "Data number", "Page number", "Paragraph number"]:
            if col not in ui_df.columns:
                ui_df[col] = None

        # ✅ PDF name = input PDF filename (first row only)
        ui_df["PDF name"] = None
        if len(ui_df) > 0:
            ui_df.iloc[0, ui_df.columns.get_loc("PDF name")] = file_name

        # ✅ Process step from UI input (all rows)
        ps = norm(str(product_name)) if product_name and str(product_name).strip() else "--"
        ui_df["Process step"] = ps

        # return exactly UI columns
        ui_df = ui_df[
            [
                "PDF name",
                "Data number",
                "Page number",
                "Paragraph number",
                "Process step",
                "Sub process step",
                "Data title",
                "Tag",
            ]
        ].copy()

        return ui_df
