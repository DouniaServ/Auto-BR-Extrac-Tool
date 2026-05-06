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
# Macro/process sections are ONLY numbered headings with a dot, e.g.
# "2. Blistering", "3. Packing - preparation...", "4. Packaging".
# TP/T codes such as "14 TP040" are subprocesses, not process steps.
MACRO_PROCESS_RE = re.compile(r"^\s*(?P<num>\d+)\.\s+(?P<title>[A-Za-z].{1,180})\s*$")
TP_CODE_PREFIX_RE = re.compile(r"^\s*\d+\s+[A-Z]{1,3}\d{2,5}\b", re.I)
STRUCTURAL_CODE_ONLY_RE = re.compile(r"^\s*\d+\s+[A-Z]{1,3}\d{2,5}(?:\s+\d{2})*(?:\s+G\d+)?\s*\.?\s*$", re.I)
PURE_NUMBER_ONLY_RE = re.compile(r"^\s*[‘'\"]?\d+(?:[.,]\d+)?[’'\"]?\s*$")
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

DASH_VAL_RE = re.compile(r"^(?P<label>.+?)\s*[−-]\s*['']?(?P<val>[^'']+)['']?\s*$")
WRAPPED_QUOTE_RE = re.compile(r"^[''\"](.*?)[''\"]$")
COLON_VAL_RE = re.compile(r"^(?P<label>.+?)\s*:\s*(?P<val>.+)$")
SPACES_SEP_RE = re.compile(r"^(?P<label>.+?)\s{2,}(?P<val>\S.+)$")

PLACEHOLDER_INT_RE = re.compile(r"^\s*\d{1,5}\s*$")

# ✅ NEW: Punctuation-only detection
PUNCT_FILLER_ONLY_RE = re.compile(
    r"""^\s*(
        [\u2022\u25CF•●\-\–\—\*_=\.·,;:|/\\]+ |
        ['''"`]+
    )\s*$""",
    re.X
)

# ✅ NEW: Units for merging
UNIT_ONLY_TOKENS = {
    "kg", "g", "mg", "µg", "ug", "ml", "l", "cl", "dl",
    "%", "ppm"
}


# ✅ NEW: split-row / continuation cleanup
CONTINUATION_START_RE = re.compile(
    r"""^\s*(
        \([^)]{1,40}\)\s*['"‘’]?\S+ |       # (number) 'HX6', (pcs) 489
        /\s*\S+ |                              # / vibrating unscrambler
        [a-z][a-z0-9,;:\- ]{1,80}$ |          # lower-case continuation line
        (and|or|of|to|for|with|from|in|on|the)\b
    )""",
    re.I | re.X,
)

SHORT_VALUE_RE = re.compile(
    r"^\s*[‘'\"]?([A-Z0-9][A-Z0-9./+\-]*|N\s*/?\s*A|ND|YES|NO|Complies|Compliant|Compatible|Labelled)[’'\"]?\s*$",
    re.I,
)

OPEN_END_RE = re.compile(r"[,/\-–—:]\s*$")


def _canon_for_dedupe(s: str) -> str:
    t = norm(s).lower()
    t = t.replace("‘", "'").replace("’", "'").replace('"', "'")
    t = re.sub(r"\s*[-–—]\s*", "-", t)
    t = re.sub(r"[^a-z0-9]+", " ", t)
    return norm(t)


def _has_unclosed_quote_or_paren(s: str) -> bool:
    t = norm(s)
    if not t:
        return False
    quote_count = sum(t.count(q) for q in ["'", "‘", "’", '"'])
    if quote_count % 2 == 1:
        return True
    return t.count("(") > t.count(")")


def _looks_like_split_continuation(title: str, prev_title: str = "") -> bool:
    """Detect rows that are fragments of the previous extracted row."""
    t = norm(title)
    p = norm(prev_title)
    if not t:
        return False

    # Examples from the output: "(number) 'HX6'", "/ vibrating unscrambler".
    if CONTINUATION_START_RE.match(t) and len(t) <= 120:
        return True

    # Short standalone values should usually attach to their label row.
    if SHORT_VALUE_RE.match(t) and len(t) <= 35:
        return True

    # Continuation after an obviously incomplete previous line.
    if p and (OPEN_END_RE.search(p) or _has_unclosed_quote_or_paren(p)):
        return True

    # Very short parenthetical/unit pieces.
    if re.fullmatch(r"\(?\s*(number|pcs|pc|kg|g|mg|ml|o?c|bar|cartons?/min|blisters?/min)\s*\)?", t, re.I):
        return True

    return False


def _merge_titles(a: str, b: str) -> str:
    a = norm(a)
    b = norm(b)
    if not a:
        return b
    if not b:
        return a
    # Avoid duplicated spaces and awkward separators.
    if b.startswith("/"):
        return norm(f"{a} {b}")
    if a.endswith("/"):
        return norm(f"{a}{b}")
    return norm(f"{a} {b}")


def merge_split_rows(rows: List[Dict]) -> List[Dict]:
    """
    Merge accidental split rows created by PDF text extraction.

    Typical fixes:
      - "- Entrance Gate" + "(number) 'HX6'" -> "- Entrance Gate (number) 'HX6'"
      - "'Tertensif SR 1,5 mg," + "loose coating'" -> one row
      - "/ vibrating unscrambler" attaches to the preceding equipment row
      - duplicate body-text rows already captured from tables are removed
    """
    merged: List[Dict] = []

    for r in rows:
        r = dict(r)
        title = norm(r.get("Data title", ""))
        value = norm(r.get("Value", ""))
        if not title and not value:
            continue

        # Delete rows that are only numbers / row indices.
        if is_pure_number_only(title) and not value:
            continue

        # Delete rows that are only structural TP codes.
        if is_structural_code_only(title) and not value:
            continue

        # Merge split fragments into previous row when context matches.
        if merged and _same_context(merged[-1], r):
            prev_title = norm(merged[-1].get("Data title", ""))

            if _looks_like_split_continuation(title, prev_title):
                merged[-1]["Data title"] = _merge_titles(prev_title, title)
                if value:
                    merged[-1]["Value"] = norm(_merge_titles(merged[-1].get("Value", ""), value))
                continue

            # Remove duplicate fragments if table extraction already gave a full row.
            cur_key = _canon_for_dedupe(title)
            if cur_key:
                for j in range(max(0, len(merged) - 8), len(merged)):
                    if not _same_context(merged[j], r):
                        continue
                    prev_key = _canon_for_dedupe(merged[j].get("Data title", ""))
                    if not prev_key:
                        continue
                    # Skip shorter duplicate/fragment rows.
                    if (cur_key in prev_key and len(cur_key) + 8 < len(prev_key)):
                        title = ""
                        break
                    if (prev_key in cur_key and len(prev_key) + 8 < len(cur_key)):
                        merged[j]["Data title"] = title
                        title = ""
                        break
                if not title:
                    continue

        r["Data title"] = title
        r["Value"] = value
        merged.append(r)

    return merged


# ============================================================
# Output cleanup / header-footer suppression
# ============================================================
# Top/bottom bands are often PDF page chrome rather than real data.
# These ratios keep the process/subprocess header band readable while
# preventing footer stamps and page banners from becoming extracted rows.
TABLE_Y_MIN_RATIO = 0.035
TABLE_Y_MAX_RATIO = 0.955
TEXT_Y_MIN_RATIO = 0.045
TEXT_Y_MAX_RATIO = 0.940

GLOBAL_HEADER_FOOTER_RE = re.compile(
    r"""
    ^\s*(AP\s+DEPARTMENT|ANPHARM|Batch\s+record\s+No|Prepared\s+by|Accepted\s*:|Approved\s+by|Product\s+code|Sample\s+doc\.|Page\s+\d+\s*/\s*\d+)\b|
    ^\s*(Download\s+the\s+packing\s+record|Return\s+of\s+packing\s+record)\b|
    ^\s*(Document\s+created\s+by|CHARGE\s+System|Generated\s+by|Created\s+by)\b|
    Date\s*:\s*\.{5,}|Hour\s*:\s*\.{5,}
    """,
    re.I | re.X,
)

GENERIC_HEADER_FOOTER_RE = re.compile(
    r"""
    ^\s*Posted\s+by\s*:|
    ^\s*Work\s+centers\s+used\s+in\s+macrophase\s*:?
    """,
    re.I | re.X,
)

# Special all-caps page titles can be valid process steps, especially page 2 OHS.
SKIP_FULL_PAGE_TITLES = set()
SPECIAL_PROCESS_TITLES = {
    "occupational safety and ergonomics": "OCCUPATIONAL SAFETY AND ERGONOMICS",
}

CHECKBOX_ONLY_RE = re.compile(r"^\s*[□☐☑☒✓✔✗✘xX\s]+\s*$")
BROKEN_PUNCT_NOISE_RE = re.compile(r"^\s*[.,;:•·_\-–—/\\|()\[\]{}'\"`´‘’“”]+\s*$")

# Remove UI/form checkbox labels that are not useful data rows, e.g.
# "Reviewed & OK checkbox", "Not OK checkbox", "OK checkbox".
CHECKBOX_LABEL_RE = re.compile(
    r"\b(reviewed\s*&?\s*ok|not\s*ok|ok)\s*(checkbox)?\b",
    re.I,
)


def is_checkbox_label(text: str) -> bool:
    t = norm(text)
    return bool(CHECKBOX_LABEL_RE.search(t))


def extract_checkbox_column(rows: List[Dict]) -> List[Dict]:
    """
    Convert checkbox label rows such as 'Reviewed & OK checkbox' into a
    real output column attached to the previous meaningful row.
    The checkbox row itself is not emitted as a Data title.
    """
    result: List[Dict] = []

    for r in rows:
        r = dict(r)
        title = norm(r.get("Data title", "") or r.get("Data Title", ""))

        if is_checkbox_label(title):
            if result:
                result[-1]["Reviewed & OK"] = "☐"
            continue

        if "Reviewed & OK" not in r:
            r["Reviewed & OK"] = ""
        result.append(r)

    return result


def is_checkbox_or_punctuation_noise(title: str) -> bool:
    t = norm(title)
    if not t:
        return True
    if CHECKBOX_ONLY_RE.match(t):
        return True
    if re.fullmatch(r"\d+\s*[□☐☑☒✓✔✗✘xX\s]+", t):
        return True
    if BROKEN_PUNCT_NOISE_RE.match(t):
        return True
    if re.fullmatch(r"[□☐☑☒✓✔✗✘xX.,;:•·_\-–—/\\|()\[\]{}'\"`´‘’“”\s]+", t):
        return True
    return False


def normalize_quote_marks(text: str) -> str:
    t = norm(text)
    t = t.replace("‘", "").replace("’", "").replace("'", "")
    t = t.replace("“", "").replace("”", "")
    return norm(t)


def normalize_subprocess_heading(text: str) -> str:
    """Normalize TP/T subprocess headings into: '<num> <code> <version> <title>'."""
    t = normalize_quote_marks(text)
    m = re.match(r"^\s*(?P<num>\d{1,2})\s+(?P<code>[A-Z]{1,3}\d{2,5})\s*(?P<rest>.*)$", t, re.I)
    if not m:
        return t
    num = m.group('num')
    code = m.group('code').upper()
    rest = norm(m.group('rest'))
    rest = re.sub(r"^\.\s*", "", rest)
    ver = ""
    title = rest

    # Standard order: "08 01 G1 . Preparation..."
    mv = re.match(r"^(?P<ver>\d{2}\s+\d{2}(?:\s+G\d+)?)\s*\.?\s*(?P<title>.*)$", rest, re.I)
    if mv:
        ver = norm(mv.group('ver').upper())
        title = norm(mv.group('title'))
    else:
        # Some PDFs return the title before the version, e.g.
        # "Preparation ... 08 01 G1 .". Move the version back after the code.
        mv_end = re.search(r"\b(?P<ver>\d{2}\s+\d{2}(?:\s+G\d+)?)\s*\.?\s*$", rest, re.I)
        if mv_end:
            ver = norm(mv_end.group('ver').upper())
            title = norm(rest[:mv_end.start()])

    title = re.sub(r"^\.\s*", "", title)
    title = re.sub(r"^(continued|continue)\s+", "", title, flags=re.I)
    return norm(" ".join(x for x in [num, code, ver, title] if x))


def fix_misordered_subprocess_heading(text: str) -> str:
    """
    Repair subprocess headings where the PDF extraction placed the first title
    word before the version, e.g.:
      '1 TP150 Cleanness 04 00 G1 . control ...'
    becomes:
      '1 TP150 04 00 G1 . Cleanness control ...'
    """
    t = norm(text)

    m = re.match(
        r"^(?P<num>\d+)\s+(?P<code>[A-Z]{1,3}\d+)\s+"
        r"(?P<title_start>[A-Za-z]+)\s+"
        r"(?P<ver>\d{2}\s+\d{2}(?:\s+G\d+)?)\s*\.?\s*"
        r"(?P<title_rest>.*)$",
        t,
        re.I,
    )

    if not m:
        return t

    return norm(
        f"{m.group('num')} {m.group('code').upper()} "
        f"{m.group('ver').upper()} . "
        f"{m.group('title_start')} {m.group('title_rest')}"
    )

# Single vertical side labels / repeated section markers leaking from margins.
MARGIN_LABELS = {"blistering", "packaging", "continued", "continue"}


FORM_COLUMN_HEADER_RE = re.compile(
    r"""
    ^\s*(Requirement|Execution\s*1?\)?|Signature|Date|Time|Hour|Date,?\s*Hour|Result\s*2?\)?|Z|NZ|ND|YES|NO|Name,?\s*ID|Name,?\s*identifier|Caption)\s*:?\s*$|
    ^\s*(Z\s+NZ|YES\s+NO|Date\s+Time|Date\s*/\s*Date\s+Time\s+Hour)\s*$|
    ^\s*(Device\s*/\s*Room\s*Name|Device/room\s*no\.?|Previous\s+product|Name,\s*Dose|Batch\s*/\s*Delivery|Number|In\s*1\)?|Pallet\s*no\.?|Code\s+number|Delivery\s+No\.?|Unit\.?\s*Measures?)\s*$|
    ^\s*(Total\s*:|TOTAL\s*:|Comments\s+on\s+the\s+process\s*:?)\s*$
    """,
    re.I | re.X,
)

SUBPROCESS_HEADER_LINE_RE = re.compile(
    rf"^\s*\d+\s+[A-Z]{{1,3}}\d{{2,5}}\b(?:\s+\d{{2}}\s+\d{{2}})?\s*\.?\s*",
    re.I,
)


def is_form_column_header_or_empty_box(title: str) -> bool:
    t = norm(title).replace("\n", " ")
    if not t:
        return True
    if FORM_COLUMN_HEADER_RE.match(t):
        return True
    # combinations made by merging adjacent table-header cells
    toks = _tokenize_simple(t)
    if toks:
        hits = sum(1 for tok in toks if tok in TABLE_HEADER_TOKENS or tok in {"yes", "no", "z", "nz", "nd"})
        if len(toks) <= 8 and hits >= max(2, int(0.6 * len(toks))):
            return True
    return False


def is_global_header_footer(title: str, process: str = "") -> bool:
    t = norm(title)
    p = norm(process)
    if not t:
        return True
    return bool(GLOBAL_HEADER_FOOTER_RE.search(t) or GLOBAL_HEADER_FOOTER_RE.search(p))


def is_generic_header_footer_or_noise(title: str, process: str = "", page_no: Optional[int] = None) -> bool:
    t = norm(title)
    p = norm(process)
    tl = t.lower().strip(" .:-")
    pl = p.lower().strip(" .:-")

    if not t:
        return True
    if is_checkbox_or_punctuation_noise(t):
        return True
    # Do not emit the process heading itself as a data row.
    if p and tl == pl:
        return True
    # Apply header/footer/page-number suppression on every page, not only page 1.
    if is_global_header_footer(t, p):
        return True
    if GENERIC_HEADER_FOOTER_RE.search(t) or GENERIC_HEADER_FOOTER_RE.search(p):
        return True
    if tl in MARGIN_LABELS and len(t.split()) <= 2:
        return True
    if pl in SKIP_FULL_PAGE_TITLES or tl in SKIP_FULL_PAGE_TITLES:
        return True
    if is_form_column_header_or_empty_box(t):
        return True
    if is_orphan_structural_row(t):
        return True
    if SUBPROCESS_HEADER_LINE_RE.match(t) and len(t.split()) <= 16:
        return True
    if is_leaked_page_header_text(t):
        return True
    if is_table_header_row([t]):
        return True
    return False


def filter_noise_records(rows: List[Dict]) -> List[Dict]:
    """Drop header/footer/page-poster records before numbering."""
    out = []
    for r in rows:
        page_no = r.get("Page number")
        title = r.get("Data title", "") or r.get("Data Title", "")
        process = r.get("Process step", "")
        if is_generic_header_footer_or_noise(title, process, page_no):
            continue
        out.append(r)
    return out


# ============================================================
# Targeted de-merge fix for accidentally flattened Page 1 rows
# ============================================================
PAGE1_ATTACHMENT_ITEMS = [
    "Empty blister",
    "PVC Film",
    "Folia AL",
    "Leaflet",
    "Cardboard",
    "Collective Label",
    "Pallet Label",
    "Production Order Summary - Print from JDEdwards",
    "Summary of the process of production and packaging of the series - v. polska",
    "Inspection of bulk product packaging",
    "Packaging Line Approval After Break",
    "Cleaning, Self-Inspection and Changeover Reports",
    "Cleanliness Inspection Reports After Changeover",
    "Registration of rejects",
]


def _copy_row_with_title(row: Dict, title: str, process: Optional[str] = None) -> Dict:
    """Create a copy of an extracted row with only Data title/process adjusted."""
    nr = dict(row)
    nr["Data title"] = normalize_quote_marks(title)
    nr["Value"] = norm(nr.get("Value", ""))
    if process is not None:
        nr["Process step"] = process
    return nr


def _split_attachment_merged_title(row: Dict) -> List[Dict]:
    """
    Split a flattened Page 1 attachment block like:
      'List of attachments Quantity Empty blister PVC Film ... Critical parameters'
    into one row per attachment item.
    """
    title = norm(row.get("Data title", ""))
    if "list of attachments" not in title.lower():
        return [row]

    found: List[Dict] = []
    low = title.lower()
    for item in PAGE1_ATTACHMENT_ITEMS:
        if item.lower() in low:
            found.append(_copy_row_with_title(row, item, process="List of attachments"))

    # If nothing matched, keep a cleaned version instead of losing data.
    if not found:
        cleaned = re.sub(r"list\s+of\s+attachments", "", title, flags=re.I)
        cleaned = re.sub(r"\bquantity\b", "", cleaned, flags=re.I)
        cleaned = re.sub(r"critical\s+parameters.*$", "", cleaned, flags=re.I)
        cleaned = norm(cleaned)
        return [_copy_row_with_title(row, cleaned, process="List of attachments")] if cleaned else []

    return found


def _split_page1_outer_cases_production_date(row: Dict) -> List[Dict]:
    """
    Split a flattened Page 1 row like:
      'number of outer cases (pcs / 488 PCS) Production date 29.10.2025'
    into:
      'number of outer cases (pcs / PCS) - 488'
      'Production date - 29.10.2025'
    """
    title = norm(row.get("Data title", ""))
    tl = title.lower()
    if not ("number of outer cases" in tl and "production date" in tl):
        return [row]

    m = re.search(
        r"(?P<label>number\s+of\s+outer\s+cases\s*\([^)]*?)\s*(?P<num>\d{1,6}(?:[,.]\d+)?)\s*(?:pcs|PCS)?\)?\s*production\s+date\s*(?P<date>\d{1,2}[./-]\d{1,2}[./-]\d{2,4})",
        title,
        re.I,
    )
    if not m:
        # Safer fallback: split at Production date without inventing structure.
        parts = re.split(r"\bProduction\s+date\b", title, maxsplit=1, flags=re.I)
        if len(parts) == 2:
            left = norm(parts[0])
            right = norm(parts[1])
            out = []
            if left:
                out.append(_copy_row_with_title(row, left))
            if right:
                out.append(_copy_row_with_title(row, f"Production date - {right}"))
            return out or [row]
        return [row]

    label = norm(m.group("label"))
    num = norm(m.group("num"))
    date = norm(m.group("date"))

    # Repair label text if the PDF put the value inside '(pcs / 488 PCS)'.
    label = re.sub(r"/\s*$", "/ PCS)", label)
    if not label.endswith(")"):
        label = label + ")"

    return [
        _copy_row_with_title(row, f"{label} - {num}"),
        _copy_row_with_title(row, f"Production date - {date}"),
    ]



def _split_page1_product_description_pairs(row: Dict) -> List[Dict]:
    """
    Split a flattened Page 1 product-description row like:
      'Number of tablets per blister/bottle - 30 Number of boxes per outer case - 120'
    into two separate rows.
    """
    title = norm(row.get("Data title", ""))
    tl = title.lower()

    if not (
        "number of tablets per blister/bottle" in tl
        and "number of boxes per outer case" in tl
    ):
        return [row]

    m = re.search(
        r"number\s+of\s+tablets\s+per\s+blister/bottle\s*[-–—]\s*(?P<t>\d+)\s+"
        r"number\s+of\s+boxes\s+per\s+outer\s+case\s*[-–—]\s*(?P<b>\d+)",
        title,
        re.I,
    )

    if not m:
        return [row]

    return [
        _copy_row_with_title(row, f"Number of tablets per blister/bottle - {m.group('t')}"),
        _copy_row_with_title(row, f"Number of boxes per outer case - {m.group('b')}"),
    ]


def _split_cleanliness_control_page3(row: Dict) -> List[Dict]:
    """
    Split a flattened Page 3 cleanliness-control block into the meaningful rows
    visible in the source table.
    """
    title = norm(row.get("Data title", ""))
    tl = title.lower()

    if not (
        "cleanliness check according to the cleaning range" in tl
        and "drum for recovered semi-finished product" in tl
    ):
        return [row]

    outputs: List[Dict] = []

    phrases = [
        "Cleanliness check according to the cleaning range for type changeover",
        '- checking the records concerning the performed activities and the type of cleaning carried out in the "Device and Room Cleaning Work Log" and in the "Cleaning, Self-Inspection and Changeover Report - Packaging Line LP1 (IMA B-44) - White Side" constituting Annex ID 0176 in the templates of ABR annexes',
        'Confirm the compliance of the cleanliness control with the provisions in the "White Side Cleanliness Check Report – Packaging Line LP1 (IMA B-44)" constituting Annex ID 0170 in the ABR Annex Templates',
        "Drum for recovered semi-finished product - OBM",
        "PT-247 container used to load semi-finished product in minibags OBM",
    ]

    # The extracted text can have small punctuation/spacing differences, so use
    # key fragments rather than only full exact-string matching.
    checks = [
        ("cleanliness check according to the cleaning range", phrases[0]),
        ("checking the records concerning the performed activities", phrases[1]),
        ("confirm the compliance of the cleanliness control", phrases[2]),
        ("drum for recovered semi-finished product", phrases[3]),
        ("pt-247 container used to load semi-finished product", phrases[4]),
    ]

    for needle, phrase in checks:
        if needle in tl:
            outputs.append(_copy_row_with_title(row, phrase))

    return outputs or [row]

def demerge_flattened_rows(rows: List[Dict]) -> List[Dict]:
    """
    Minimal targeted fix only: do not change the extractor logic; just split
    known flattened rows caused by PDF table parsing.
    """
    out: List[Dict] = []
    seen = set()

    for r in rows:
        expanded = [r]

        tmp: List[Dict] = []
        for x in expanded:
            tmp.extend(_split_page1_outer_cases_production_date(x))

        expanded = []
        for x in tmp:
            expanded.extend(_split_page1_product_description_pairs(x))

        tmp = []
        for x in expanded:
            tmp.extend(_split_cleanliness_control_page3(x))

        expanded = []
        for x in tmp:
            expanded.extend(_split_attachment_merged_title(x))

        for x in expanded:
            title = norm(x.get("Data title", ""))
            if not title:
                continue
            key = (
                x.get("Page number"),
                norm(x.get("Process step", "")).lower(),
                norm(x.get("Sub Process step", "")).lower(),
                title.lower(),
            )
            if key in seen:
                continue
            seen.add(key)
            out.append(x)

    return out


# ============================================================
# Header detection
# ============================================================
_PL_MAP = str.maketrans({
    "ą": "a", "ć": "c", "ę": "e", "ł": "l", "ń": "n", "ó": "o", "ś": "s", "ż": "z", "ź": "z",
    "Ą": "a", "Ć": "c", "Ę": "e", "Ł": "l", "Ń": "n", "Ó": "o", "Ś": "s", "Ż": "z", "Ź": "z",
})


def norm(s: str) -> str:
    return WS_RE.sub(" ", str(s or "").replace("\u00A0", " ")).strip()


def is_structural_code_only(text: str) -> bool:
    """True for standalone subprocess markers like '14 TP040' or '14 TP040 03 38'."""
    return bool(STRUCTURAL_CODE_ONLY_RE.match(norm(text)))


def is_pure_number_only(text: str) -> bool:
    """True for orphan rows containing only a number, e.g. pallet row markers 1, 2, 3."""
    return bool(PURE_NUMBER_ONLY_RE.match(norm(text)))


def is_orphan_structural_row(text: str, value: str = "") -> bool:
    t = norm(text)
    v = norm(value)
    if not t and not v:
        return True
    if v:
        return False
    if is_structural_code_only(t):
        return True
    if is_pure_number_only(t):
        return True
    return False


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
    # IMPORTANT for Windows: use a context manager so the PDF file handle
    # is released before TemporaryDirectory tries to delete input.pdf.
    with fitz.open(pdf_path) as doc:
        page = doc[0]
        words = page.get_text("words") or []


    lines_map: Dict[Tuple[int, int], List[tuple]] = {}
    for w in words:
        if len(w) < 8:
            continue
        x0, y0, x1, y1, text, bno, lno, wno = w
        text = norm(text)
        if not text:
            continue
        lines_map.setdefault((int(bno), int(lno)), []).append((float(x0), float(x1), float(y0), text))

    ordered = []
    for (bno, lno), items in lines_map.items():
        items.sort(key=lambda t: t[0])
        line_text = norm(" ".join(t[3] for t in items))
        if not line_text:
            continue
        y = min(t[2] for t in items)
        ordered.append((y, bno, lno, line_text))
    ordered.sort(key=lambda t: (t[0], t[1], t[2]))

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
            "Data title": norm(f"{label} - {val}") if val else label,
            "Value": "",
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


def detect_special_process_from_band(lines: List[Dict]) -> Tuple[str, Optional[int], float]:
    """Detect valid non-numbered process headings such as page 2 OHS."""
    for ln in lines:
        t = norm(ln.get("text", ""))
        key = t.lower().strip(" .:-")
        if key in SPECIAL_PROCESS_TITLES:
            return SPECIAL_PROCESS_TITLES[key], None, float(ln.get("max_size", 0.0) or 0.0)
    return "", None, 0.0


def detect_process_step_from_band(
    lines: List[Dict],
    prev_process_num: Optional[int],
    prev_process_size: float,
) -> Tuple[str, Optional[int], float]:
    """
    Detect ONLY macro process sections.

    Correct hierarchy for Anpharm PDFs:
      Process step      = high-level section, e.g. "4. Packaging"
      Sub Process step  = procedure/T-code, e.g. "14 TP040 03 38 ..."

    Older logic treated all-caps / short T-code headers as Process step, which
    caused repeated values such as "14 TP040", "15 TP020", "16 TP020" in the
    Process step column. This function intentionally ignores T-code headers.
    """
    candidates = []
    for ln in lines:
        t = norm(ln.get("text", ""))
        if not t:
            continue
        if HEADER_BLOCK_RE.search(t):
            continue
        if TP_CODE_PREFIX_RE.match(t):
            continue
        if is_form_column_header_or_empty_box(t):
            continue
        if is_global_header_footer(t):
            continue

        m = MACRO_PROCESS_RE.match(t)
        if not m:
            continue

        title = norm(t)
        # Avoid false positives such as "1. TABLETS/CAPSULES..." inside tables.
        # Real section headings are short and appear near the top/left.
        if len(title) > 120:
            continue
        try:
            num = int(m.group("num"))
        except Exception:
            continue
        candidates.append((ln.get("max_size", 0.0), ln.get("y", 0.0), title, num))

    if not candidates:
        return "", None, 0.0

    candidates.sort(key=lambda x: (-float(x[0] or 0), float(x[1] or 0)))
    best_size, _, best_text, best_num = candidates[0]

    # Do not move backwards to a smaller process number unless the heading is
    # clearly at least as prominent as the previous detected section.
    if prev_process_num is not None and best_num < prev_process_num:
        if float(best_size or 0) + 0.2 < float(prev_process_size or 0):
            return "", None, 0.0

    return best_text, best_num, float(best_size or 0.0)


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
            return normalize_subprocess_heading(_clean_rotated_append(t))

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
        return normalize_subprocess_heading(_clean_rotated_append(norm(prefix + (" " + title if title else ""))))

    return ""


# ============================================================
# Page 1 — tables (✅ IMPROVED: Better multi-cell merging)
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
        # Crop away page chrome. This prevents page 1 top banner/footer
        # from being interpreted as normal table rows.
        w, h = page.width, page.height
        page_cropped = page.crop((0, h * TABLE_Y_MIN_RATIO, w, h * TABLE_Y_MAX_RATIO))
        tables = page_cropped.extract_tables() or []

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
                merged = norm(" ".join(header_non_empty))

                records.append({
                    "Page number": 1,
                    "Paragraph number": paragraph_table_no,
                    "Process step": process,
                    "Sub Process step": "",
                    "Data title": merged,
                })
                first_row_emitted = True

            elif len(header_non_empty) == 1:
                lbl_parsed, val_parsed = parse_label_value_from_line(header_non_empty[0])
                if val_parsed and _looks_like_label(lbl_parsed) and _looks_like_value(val_parsed):
                    merged = norm(f"{lbl_parsed} - {val_parsed}")
                    records.append({
                        "Page number": 1,
                        "Paragraph number": paragraph_table_no,
                        "Process step": process,
                        "Sub Process step": "",
                        "Data title": merged,
                        "Value": "",
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

                # ✅ IMPROVED: Merge all cells into single data title
                left_parts = non_empty
                merged = norm(" ".join(left_parts))

                if not merged or HEADER_BLOCK_RE.search(merged):
                    continue

                lbl2, inline_val = split_inline_dash_value(merged)
                if inline_val:
                    merged = lbl2 + " - " + inline_val

                records.append({
                    "Page number": 1,
                    "Paragraph number": paragraph_table_no,
                    "Process step": process,
                    "Sub Process step": "",
                    "Data title": merged,
                })

    return records


# ============================================================
# Tables on pages 2+ (✅ IMPROVED: Better cell merging)
# ============================================================
def build_label_value_from_cells(non_empty: List[str]) -> Tuple[str, str]:
    """✅ IMPROVED: Merge all cells into Data title (Value column empty)"""
    if not non_empty:
        return "", ""
    
    # ✅ Merge ALL cells together into Data title
    merged = norm(" ".join(non_empty))
    merged = unquote_if_wrapped(merged)

    return merged, ""


def extract_tables_with_subprocess_from_row(
    pdf_path: str,
    page_index0: int,
    current_process: str,
    current_sub: str
) -> List[Dict]:
    rows_out: List[Dict] = []

    with pdfplumber.open(pdf_path) as pdf:
        page = pdf.pages[page_index0]
        w, h = page.width, page.height
        page_cropped = page.crop((0, h * TABLE_Y_MIN_RATIO, w, h * TABLE_Y_MAX_RATIO))
        tables = page_cropped.extract_tables() or []

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
                if is_checkbox_or_punctuation_noise(label):
                    continue
                label = normalize_quote_marks(label)

                if is_leaked_page_header_text(label):
                    continue

                if looks_like_instruction_row(label):
                    continue

                # Skip orphan structural rows such as '14 TP040', '14 TP040 03 38', or bare pallet numbers '1'.
                if is_orphan_structural_row(label, value):
                    continue

                rows_out.append({
                    "Page number": page_index0 + 1,
                    "Paragraph number": None,
                    "Process step": current_process,
                    "Sub Process step": current_sub,
                    "Data title": label,
                })

    return rows_out


# ============================================================
# Pages 2+ extraction
# ============================================================
def extract_pages_2plus_target(pdf_path: str, start_page: int = 2) -> List[Dict]:
    out: List[Dict] = []

    current_process = ""
    current_process_num: Optional[int] = None
    current_process_size: float = 0.0
    current_sub = ""

    # IMPORTANT for Windows: keep the PyMuPDF document open only inside
    # this block, then release it before TemporaryDirectory cleanup.
    with fitz.open(pdf_path) as doc:
        for pno0 in range(start_page - 1, len(doc)):
            page = doc[pno0]
            page_no = pno0 + 1

            header_lines = page_lines_grouped(page, y_min_ratio=0.00, y_max_ratio=0.35, y_tol=2.6)
            body_lines = page_lines_grouped(page, y_min_ratio=TEXT_Y_MIN_RATIO, y_max_ratio=TEXT_Y_MAX_RATIO, y_tol=2.6)

            proc, pnum, psize = detect_process_step_from_band(
                header_lines,
                prev_process_num=current_process_num,
                prev_process_size=current_process_size
            )
            if not proc:
                proc, pnum, psize = detect_special_process_from_band(header_lines + body_lines[:8])
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
                key = (r.get("Process step", ""), r.get("Sub Process step", ""), r.get("Data title", ""), r.get("Value", ""))
                if key in seen:
                    continue
                seen.add(key)
                cleaned_table_rows.append(r)

            for r in cleaned_table_rows:
                r["Paragraph number"] = paragraph_no
                out.append(r)

            existing_titles = set(r["Data title"] for r in cleaned_table_rows)

            # Avoid duplicate raw PyMuPDF body lines when pdfplumber already
            # extracted structured rows from the visible table on this page.
            if len(cleaned_table_rows) >= 2:
                continue

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

                # ✅ INTEGRATE: Merge value into data title
                if val:
                    label = norm(f"{label} - {val}")

                if (not val) and label.rstrip().endswith(":") and i < len(body_lines):
                    nxt = body_lines[i]["text"]
                    if nxt:
                        if not is_leaked_page_header_text(nxt) and not looks_like_instruction_row(nxt):
                            if not re.match(r"^\s*\d+\s*/\s*\d+\s*$", nxt) and not MANY_SPACES_RE.search(nxt):
                                val_nxt = unquote_if_wrapped(nxt)
                                label = norm(f"{label} - {val_nxt}")
                                i += 1

                if is_leaked_page_header_text(label):
                    continue
                if is_checkbox_or_punctuation_noise(label):
                    continue
                label = normalize_quote_marks(label)

                # Skip orphan structural/body rows such as '14 TP040' or bare numbers.
                if is_orphan_structural_row(label, ""):
                    continue

                out.append({
                    "Page number": page_no,
                    "Paragraph number": paragraph_no,
                    "Process step": current_process,
                    "Sub Process step": current_sub,
                    "Data title": label,
                })

    return out

# ============================================================
# ✅ NEW: Cleanup functions for fragment merging
# ============================================================
def _is_punct_or_filler_only(title: str) -> bool:
    """Check if title contains only punctuation/filler"""
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
    """Check if title is a single word (potential fragment)"""
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
    """Determine if this row should be merged into previous row's title"""
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
    """Check if two rows share the same context (page, paragraph, process, subprocess)"""
    keys = ["Page number", "Paragraph number", "Process step", "Sub Process step"]
    return all(norm(prev.get(k, "")) == norm(cur.get(k, "")) for k in keys)


def clean_skip_and_merge_fragments(rows: List[Dict]) -> List[Dict]:
    """
    ✅ IMPROVED: Three-stage cleanup
    1. Skip rows with only punctuation/filler (no value)
    2. Merge single-word fragments into previous row's Data title
    3. Merge unit-only tokens into previous row's Data title
    """
    cleaned: List[Dict] = []

    for r in rows:
        title = normalize_quote_marks(r.get("Data title", ""))
        value = normalize_quote_marks(r.get("Value", ""))

        if is_checkbox_or_punctuation_noise(title) and not value:
            continue

        # Stage 0: Skip page headers, footers, page posters and margin labels
        if is_generic_header_footer_or_noise(title, r.get("Process step", ""), r.get("Page number")):
            continue

        # Stage 1: Skip pure filler rows
        if _is_punct_or_filler_only(title) and not value:
            continue

        # Stage 2: Merge fragments into previous row
        if cleaned and _should_merge_into_previous(title, value) and _same_context(cleaned[-1], r):
            prev = cleaned[-1]
            prev_title = norm(prev.get("Data title", ""))

            fragment = title
            if value:
                fragment = norm(f"{title} {value}") if title else value

            if fragment:
                prev["Data title"] = norm(f"{prev_title} {fragment}") if prev_title else fragment

            continue

        r["Data title"] = title
        r["Value"] = value
        cleaned.append(r)

    return cleaned


def clean_process_hierarchy(rows: List[Dict]) -> List[Dict]:
    """
    Enforce the final hierarchy:
      - Process step is only the macro section: "1. ...", "2. ...", "3. ..."
      - Sub Process step carries TP/T-code lines: "14 TP040 ..."

    This also repairs any rows already contaminated by a T-code in Process step
    by carrying forward the last valid macro process.
    """
    cleaned: List[Dict] = []
    current_macro = ""

    for r in rows:
        r = dict(r)
        proc = norm(r.get("Process step", ""))
        sub = fix_misordered_subprocess_heading(
            normalize_subprocess_heading(r.get("Sub Process step", ""))
        )
        title = normalize_quote_marks(r.get("Data title", ""))

        # Capture true macro process headings.
        if MACRO_PROCESS_RE.match(proc) and not TP_CODE_PREFIX_RE.match(proc):
            current_macro = proc
        elif MACRO_PROCESS_RE.match(title) and not TP_CODE_PREFIX_RE.match(title):
            current_macro = title

        # If Process step is actually a TP/T-code, move/preserve it as subprocess.
        if TP_CODE_PREFIX_RE.match(proc):
            if not sub or sub == "--":
                sub = proc
            proc = current_macro

        # If the Data title is only a short subprocess marker, capture it as
        # context and do not emit a data row.
        if is_structural_code_only(title):
            if not sub or sub == "--":
                sub = title
            current_sub = sub
            continue

        # If no macro is available yet, keep harmless page-1 table process labels
        # such as "Product description" / "Batch size". For pages with T-codes,
        # the current macro will be set after the first section heading.
        if current_macro and (not proc or TP_CODE_PREFIX_RE.match(proc)):
            proc = current_macro

        # A full TP/T-code heading in Data title is subprocess context, not data.
        if TP_CODE_PREFIX_RE.match(title):
            if not sub or sub == "--" or len(sub) < len(title):
                sub = title
            continue

        # Remove orphan numeric rows.
        if is_pure_number_only(title):
            continue
        if is_checkbox_or_punctuation_noise(title):
            continue
        if is_global_header_footer(title, proc):
            continue

        r["Process step"] = proc
        r["Sub Process step"] = sub
        r["Data title"] = title
        cleaned.append(r)

    return cleaned


# ============================================================
# Tag + review (✅ updated compute function)
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
# Main export (✅ INCLUDES cleanup pipeline)
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

    # ✅ Clean output rows (skip punctuation-only/header-footer + merge fragments)
    all_rows = clean_skip_and_merge_fragments(page1 + p2plus)
    all_rows = clean_process_hierarchy(all_rows)
    all_rows = merge_split_rows(all_rows)
    all_rows = filter_noise_records(all_rows)
    all_rows = demerge_flattened_rows(all_rows)
    all_rows = extract_checkbox_column(all_rows)

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
        "Reviewed & OK",
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
# UI extract wrapper (✅ INCLUDES cleanup)
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

        # --- SAME extraction as original code ---
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

        # ✅ Clean output rows (skip punctuation-only/header-footer + merge fragments)
        all_rows = clean_skip_and_merge_fragments(page1 + p2plus)
        all_rows = clean_process_hierarchy(all_rows)
        all_rows = merge_split_rows(all_rows)
        all_rows = filter_noise_records(all_rows)
        all_rows = demerge_flattened_rows(all_rows)
        all_rows = extract_checkbox_column(all_rows)

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
            "Reviewed & OK",
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
        ui_df["Reviewed & OK"] = ui_df["Reviewed & OK"].fillna("").astype(str)
        ui_df["Value"] = ui_df["Value"].fillna("").astype(str)

        return ui_df[required].copy()
