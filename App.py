from __future__ import annotations

import os
import base64
import io
import time
import uuid
import hashlib
import traceback
from dataclasses import dataclass
from typing import List, Optional, Tuple, Dict, Any

import pandas as pd
import streamlit as st
import openpyxl  # needed to preserve VBA from template.xlsm
from openpyxl.styles import Alignment, Font

# Optional AG Grid
HAS_AGGRID = False
try:
    from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode, DataReturnMode
    HAS_AGGRID = True
except Exception:
    HAS_AGGRID = False


# =============================================================================
# Tool version (V1.0)
# =============================================================================
TOOL_VERSION = "1.0.0"

# =============================================================================
# Extractor loader (keeps each site's logic untouched)
# =============================================================================
def get_extractor(site: str):
    if site == "Arklow":
        from sites import arklow
        return arklow.extract
    if site == "Anpharm":
        from sites import anpharm
        return anpharm.extract
    if site == "Gidy":
        from sites import gidy
        return gidy.extract
    if site == "Toledo":
        from sites import toledo
        return toledo.extract
    if site == "Bolbec":
        from sites import bolbec
        return bolbec.extract
    raise ValueError(f"No extractor implemented for site: {site}")


# =============================================================================
# Per-site schema
# =============================================================================
SITES = ["Arklow", "Anpharm", "Gidy", "Toledo", "Bolbec"]

REVIEW_OK_COL = "Reviewed & OK"
REVIEWER_COL = "Reviewer"
REVIEWED_AT_COL = "Reviewed at"
REVIEWER_NOTE_COL = "Reviewer note"

SITE_SCHEMA: Dict[str, Dict[str, Any]] = {
    "Arklow": {
        "columns": [
            "PDF name",
            "Data number",
            "Page number",
            "Paragraph number",
            "Process step",
            "Sub process step",
            "Data title",
            "Tag",
            REVIEW_OK_COL,
            REVIEWER_COL,
            REVIEWED_AT_COL,
            REVIEWER_NOTE_COL,
        ],
        "required_nonempty": ["Data number", "Page number", "Process step", "Data title"],
        "page_col": "Page number",
        "page_min": 1,
        "dup_key": ["Page number", "Paragraph number", "Data title", "Tag"],
        "search_cols": ["Process step", "Sub process step", "Data title", "Tag"],
        "review_ok_col": REVIEW_OK_COL,
        "reviewer_col": REVIEWER_COL,
        "reviewed_at_col": REVIEWED_AT_COL,
        "review_note_col": REVIEWER_NOTE_COL,
    },
    "Anpharm": {
        "columns": [
            "PDF name",
            "Data number",
            "Page number",
            "Paragraph number",
            "Process step",
            "Sub process step",
            "Data Title",
            "Data tag",
            REVIEW_OK_COL,
            REVIEWER_COL,
            REVIEWED_AT_COL,
            REVIEWER_NOTE_COL,
        ],
        "required_nonempty": ["Data number", "Page number", "Process step", "Data Title"],
        "page_col": "Page number",
        "page_min": 1,
        "dup_key": ["Page number", "Paragraph number", "Data Title", "Value", "Data tag"],
        "search_cols": ["Process step", "Sub process step", "Data Title", "Value", "Data tag"],
        "review_ok_col": REVIEW_OK_COL,
        "reviewer_col": REVIEWER_COL,
        "reviewed_at_col": REVIEWED_AT_COL,
        "review_note_col": REVIEWER_NOTE_COL,
    },
    "Gidy": {
        "columns": [
            "PDF name",
            "Data number",
            "Page number",
            "Paragraph number",
            "Process step",
            "Sub process step",
            "Data Title",
            "Data tag",
            REVIEW_OK_COL,
            REVIEWER_COL,
            REVIEWED_AT_COL,
            REVIEWER_NOTE_COL,
        ],
        "required_nonempty": ["Data number", "Page number", "Process step", "Data Title"],
        "page_col": "Page number",
        "page_min": 1,
        "dup_key": ["Page number", "Paragraph number", "Data Title", "Data tag"],
        "search_cols": ["Process step", "Sub process step", "Data Title", "Data tag", "Théorique", "Réel", "Visa"],
        "review_ok_col": REVIEW_OK_COL,
        "reviewer_col": REVIEWER_COL,
        "reviewed_at_col": REVIEWED_AT_COL,
        "review_note_col": REVIEWER_NOTE_COL,
    },
    "Toledo": {
        "columns": [],
        "required_nonempty": [],
        "page_col": None,
        "page_min": 1,
        "dup_key": [],
        "search_cols": [],
        "review_ok_col": REVIEW_OK_COL,
        "reviewer_col": REVIEWER_COL,
        "reviewed_at_col": REVIEWED_AT_COL,
        "review_note_col": REVIEWER_NOTE_COL,
    },
    "Bolbec": {
        "columns": [],
        "required_nonempty": [],
        "page_col": None,
        "page_min": 1,
        "dup_key": [],
        "search_cols": [],
        "review_ok_col": REVIEW_OK_COL,
        "reviewer_col": REVIEWER_COL,
        "reviewed_at_col": REVIEWED_AT_COL,
        "review_note_col": REVIEWER_NOTE_COL,
    },
}

REVIEW_STATUS_CHOICES = ["Draft", "In review", "Approved", "Rejected"]


# =============================================================================
# UI Config + CSS
# =============================================================================
st.set_page_config(page_title="Auto BR Extractor", page_icon="📄", layout="wide")


def inject_background_css(image_name="background image 2.jpg"):
    paths = [
        image_name,
        os.path.join(os.getcwd(), image_name),
        os.path.join(os.path.dirname(__file__), image_name),
    ]
    real_path = next((p for p in paths if os.path.exists(p)), None)

    if real_path:
        with open(real_path, "rb") as f:
            b64 = base64.b64encode(f.read()).decode()
        ext = os.path.splitext(real_path)[1].lower()
        mime = "png" if ext == ".png" else "jpeg"
        bg = f'url("data:image/{mime};base64,{b64}")'
    else:
        bg = "linear-gradient(135deg, #020b1f, #08162b)"

    st.markdown(
        f"""
        <style>
        html, body, .stApp,
        [data-testid="stAppViewContainer"],
        [data-testid="stAppViewContainer"] > .main {{
            background: {bg} !important;
            background-size: cover !important;
            background-position: center center !important;
            background-repeat: no-repeat !important;
            background-attachment: fixed !important;
        }}
        section.main {{ background: transparent !important; }}
        </style>
        """,
        unsafe_allow_html=True,
    )


def inject_ui_css():
    st.markdown(
        """
        <style>
        html, body, [class*="css"], .stMarkdown, .stText, .stMarkdown p {
          color: #ffffff !important;
          font-family: "Inter", Arial, sans-serif !important;
        }
        .stTabs [data-baseweb="tab"] {
          font-family: Arial, sans-serif !important;
          font-weight: 900 !important;
          color: #ffffff !important;
          font-size: 18px !important;
        }
        .stTabs [data-baseweb="tab"][aria-selected="true"] { color: #ffffff !important; }

        div[data-testid="stWidgetLabel"],
        div[data-testid="stWidgetLabel"] *,
        label,
        .stTextInput label,
        .stTextArea label,
        .stSelectbox label,
        .stNumberInput label,
        .stFileUploader label,
        .stMultiSelect label,
        .stRadio label,
        .stCheckbox label,
        [data-baseweb="form-control-label"],
        [data-baseweb="form-control-label"] * {
          color: #ffffff !important;
          opacity: 1 !important;
          font-weight: 900 !important;
        }

        h3,
        div[data-testid="stSubheader"],
        div[data-testid="stSubheader"] *,
        h2 {
          color: #ffffff !important;
          font-weight: 1000 !important;
          opacity: 1 !important;
        }

        input, textarea { color: #08162b !important; font-weight: 900 !important; }

        .stCaption, div[data-testid="stCaptionContainer"] * {
          color: #ffffff !important;
          opacity: 0.95 !important;
          font-weight: 800 !important;
        }

        .kebab-btn button{
          border-radius: 12px !important;
          padding: 6px 12px !important;
          font-size: 22px !important;
          font-weight: 900 !important;
          background: rgba(255,255,255,0.14) !important;
          border: 1px solid rgba(255,255,255,0.35) !important;
          color: #ffffff !important;
          box-shadow: 0 10px 24px rgba(0,0,0,0.35) !important;
        }
        .kebab-btn button:hover{ background: rgba(255,255,255,0.22) !important; }

        div[data-testid="stExpander"]{
          background: rgba(10, 18, 35, 0.62) !important;
          border: 1px solid rgba(255,255,255,0.18) !important;
          border-radius: 16px !important;
          box-shadow: 0 10px 28px rgba(0,0,0,0.35) !important;
          backdrop-filter: blur(10px) !important;
        }
        div[data-testid="stExpander"] *{ color: #ffffff !important; font-weight: 800 !important; }

        .metric-card{
          background: rgba(10, 18, 35, 0.62);
          border-radius: 16px;
          padding: 18px;
          border: 1px solid rgba(255,255,255,0.18);
          box-shadow: 0 10px 28px rgba(0,0,0,0.35);
          backdrop-filter: blur(10px);
        }
        .metric-label{
          font-size: 13px;
          font-weight: 900;
          letter-spacing: 0.3px;
          color: #ffffff !important;
          margin-bottom: 6px;
          text-transform: uppercase;
        }
        .metric-value{
          font-size: 22px;
          font-weight: 900;
          color: #ffffff !important;
          word-break: break-word;
        }
        .hr{ height: 1px; background: rgba(255,255,255,0.22); margin: 12px 0 14px; }

        div[data-baseweb="select"] > div{
          background: #ffffff !important;
          border: 2px solid rgba(255,255,255,0.92) !important;
          border-radius: 14px !important;
          min-height: 54px !important;
          box-shadow: 0 10px 22px rgba(0,0,0,0.25) !important;
        }
        div[data-baseweb="select"] span{
          color: #08162b !important;
          font-size: 18px !important;
          font-weight: 900 !important;
        }
        div[data-baseweb="select"] svg{ color: #08162b !important; }
        div[data-baseweb="menu"]{ background: #ffffff !important; border-radius: 14px !important; }

        .panel{
          background: rgba(10, 18, 35, 0.62);
          border: 1px solid rgba(255,255,255,0.18);
          border-radius: 16px;
          padding: 16px;
          box-shadow: 0 10px 28px rgba(0,0,0,0.35);
          backdrop-filter: blur(10px);
        }

        div[data-testid="stDataFrame"],
        div[data-testid="stDataFrame"] * { color: #08162b !important; }

        .footer-version{
          position: fixed;
          bottom: 14px;
          left: 0;
          width: 100%;
          text-align: center;
          opacity: 0.95;
          font-size: 20px;
          font-weight: 1000;
          letter-spacing: 0.5px;
          color: #ffffff;
          z-index: 9999;
          pointer-events: none;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


def step_title(step_key: str, text: str):
    active = st.session_state.view_step == step_key
    color = "#ffffff" if active else "rgba(255,255,255,0.75)"
    st.markdown(
        f"<h2 style='color:{color}; font-weight:1000; margin-bottom:16px;'>{text}</h2>",
        unsafe_allow_html=True,
    )


inject_background_css("background image 2.jpg")
inject_ui_css()


# =============================================================================
# NEW: Arklow-only metadata fill (Product name -> Process step)
# =============================================================================
def apply_ui_metadata_to_df(site: str, df: pd.DataFrame, file_name: str) -> pd.DataFrame:
    """
    Sites using UI Product name for Process step:
      - Arklow
      - Bolbec
      - Toledo

    Also sets file name in first row if column exists.
    """
    if not isinstance(df, pd.DataFrame) or df.empty:
        return df

    # Only apply to selected sites
    if site not in {"Arklow", "Bolbec", "Toledo"}:
        return df

    product = (st.session_state.product_name or "").strip() or "--"

    # Process step override
    if "Process step" in df.columns:
        df["Process step"] = product

    # File name in first row
    if "Unnamed: 0" in df.columns:
        df["Unnamed: 0"] = None
        df.loc[df.index[0], "Unnamed: 0"] = file_name

    if "PDF name" in df.columns:
        df["PDF name"] = None
        df.loc[df.index[0], "PDF name"] = file_name

    return df


# =============================================================================
# State helpers
# =============================================================================
def get_active_schema(site: str) -> Dict[str, Any]:
    return SITE_SCHEMA.get(site, {}) or {}


def derive_columns_from_schema_or_df(df: Optional[pd.DataFrame], schema: Dict[str, Any]) -> List[str]:
    cols = schema.get("columns") or []
    if cols:
        return cols
    if isinstance(df, pd.DataFrame):
        return list(df.columns)
    return []


def derive_search_cols(schema: Dict[str, Any], cols: List[str]) -> List[str]:
    sc = schema.get("search_cols") or []
    if sc:
        return [c for c in sc if c in cols]
    return cols


def ss_init():
    ss = st.session_state
    ss.setdefault("job_id", str(uuid.uuid4())[:8])
    ss.setdefault("site", SITES[0])
    ss.setdefault("product_name", "")

    ss.setdefault("batch_record_bytes", None)
    ss.setdefault("batch_record_name", None)
    ss.setdefault("batch_record_fp", None)

    ss.setdefault("df", None)

    # Timing / KPI
    ss.setdefault("extraction_started_at", "")
    ss.setdefault("extraction_ended_at", "")
    ss.setdefault("extraction_seconds", None)      # float seconds
    ss.setdefault("extraction_rows", 0)           # int
    ss.setdefault("extraction_source", "")        # "pdf_word" | "review_pack"

    # Validation (user-action only)
    ss.setdefault("validation", None)
    ss.setdefault("user_validated", False)
    ss.setdefault("validation_status", "Not validated yet")

    # keep only search + view_mode
    ss.setdefault("search", "")
    ss.setdefault("view_mode", "Not reviewed only")

    ss.setdefault("audit_log", [])
    ss.setdefault("show_menu", False)

    ss.setdefault("reviewer_name", os.getenv("REVIEWER_DEFAULT", ""))
    ss.setdefault("review_status", "Draft")
    ss.setdefault("review_comment", "")
    ss.setdefault("reviewed_at", "")
    ss.setdefault("baseline_hash", "")

    ss.setdefault("skipped_pages", [])
    ss.setdefault("view_step", "Setup")

    ss.setdefault("blocking_msgs", [])

    ss.setdefault("active_schema", get_active_schema(ss.site))
    ss.setdefault("active_columns", derive_columns_from_schema_or_df(ss.df, ss.active_schema))
    ss.setdefault("extraction_control_df", pd.DataFrame())


def audit(event: str, details: str = ""):
    st.session_state.audit_log.append(
        {
            "ts": time.strftime("%Y-%m-%d %H:%M:%S"),
            "job_id": st.session_state.job_id,
            "event": event,
            "details": details,
        }
    )


def reset_all():
    st.session_state.clear()
    ss_init()


ss_init()


# =============================================================================
# Guards (steps always enabled; clicking warns)
# =============================================================================
def needs_product() -> bool:
    return not bool((st.session_state.product_name or "").strip())


def needs_data() -> bool:
    return not isinstance(st.session_state.df, pd.DataFrame)


def set_blocking_messages(msgs: List[str]):
    st.session_state.blocking_msgs = msgs


def go_or_warn(target_step: str):
    msgs = []
    if needs_product():
        msgs.append("• Please enter **Product name** first (Step 1).")
    if needs_data():
        msgs.append("• Please **upload Batch Record (PDF/Word)** OR **upload Review Pack (Excel)** first (Step 1).")

    if msgs:
        set_blocking_messages(msgs)
        st.session_state.view_step = "Setup"
        return

    set_blocking_messages([])
    st.session_state.view_step = target_step


def enforce_step_guard_soft():
    """
    Safety net if someone ends up in Review/Export without prerequisites.
    """
    if st.session_state.view_step in ("Review", "Export Excel"):
        msgs = []
        if needs_product():
            msgs.append("• Please enter **Product name** first (Step 1).")
        if needs_data():
            msgs.append("• Please **upload Batch Record (PDF/Word)** OR **upload Review Pack (Excel)** first (Step 1).")

        if msgs:
            set_blocking_messages(msgs)
            st.session_state.view_step = "Setup"
            st.stop()

def style_checklist_column(ws, header_name=REVIEW_OK_COL):
    checklist_col = None

    for cell in ws[1]:
        if str(cell.value).strip() == header_name:
            checklist_col = cell.column
            break

    if checklist_col is None:
        return

    for row in range(2, ws.max_row + 1):
        cell = ws.cell(row=row, column=checklist_col)

        value = str(cell.value).strip().lower()

        if value in {"true", "1", "yes", "y", "checked", "☑", "☒", "✓", "✔", "x", "þ"}:
            cell.value = "þ"   # checked checkbox in Wingdings
        else:
            cell.value = "¨"   # unchecked checkbox in Wingdings

        cell.font = Font(name="Wingdings", size=16, bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center")
# =============================================================================
# Hash + dirty
# =============================================================================
def df_content_hash(df: pd.DataFrame) -> str:
    b = df.to_csv(index=False).encode("utf-8")
    return hashlib.sha256(b).hexdigest()[:16]


def mark_dirty(reason: str):
    st.session_state.user_validated = False
    st.session_state.validation_status = "Not validated yet"

    if st.session_state.review_status == "Approved":
        st.session_state.review_status = "Draft"
        st.session_state.reviewed_at = ""
        audit("approval_revoked_on_edit", reason)


# =============================================================================
# Validation
# =============================================================================
@dataclass
class ValidationResult:
    ok: bool
    errors: pd.DataFrame
    warnings: pd.DataFrame
    flagged_row_ids: List[int]


def validate_df(df: pd.DataFrame, schema: Dict[str, Any]) -> ValidationResult:
    cols = derive_columns_from_schema_or_df(df, schema)
    required = schema.get("required_nonempty") or []
    page_col = schema.get("page_col")
    page_min = int(schema.get("page_min", 1))
    dup_key = schema.get("dup_key") or []

    errors: List[dict] = []
    warnings: List[dict] = []
    flagged = set()

    # schema cols missing
    missing_cols = [c for c in cols if c not in df.columns]
    if missing_cols:
        for c in missing_cols:
            errors.append({"type": "schema", "row": None, "column": c, "message": f"Missing column: {c}"})
        return ValidationResult(False, pd.DataFrame(errors), pd.DataFrame(warnings), [])

    # required non-empty
    for col in required:
        if col not in df.columns:
            errors.append({"type": "schema", "row": None, "column": col, "message": f"Missing required column: {col}"})
            continue
        empty = df[col].isna() | (df[col].astype(str).str.strip() == "")
        for i in df.index[empty].tolist()[:5000]:
            errors.append({"type": "required", "row": int(i), "column": col, "message": "Value required"})
            flagged.add(int(i))

    # page checks
    if page_col and page_col in df.columns:
        p = pd.to_numeric(df[page_col], errors="coerce")
        bad_page = p.isna() | (p < page_min) | (p % 1 != 0)
        for i in df.index[bad_page].tolist()[:5000]:
            warnings.append(
                {"type": "page", "row": int(i), "column": page_col, "message": f"Page must be integer >= {page_min}"}
            )
            flagged.add(int(i))

    # duplicates (warning) ✅ V1.0 implemented
    if dup_key:
        usable = [c for c in dup_key if c in df.columns]
        if usable:
            key_df = df[usable].copy()

            nonempty_mask = True
            for c in usable:
                nonempty_mask = nonempty_mask & (~key_df[c].isna()) & (key_df[c].astype(str).str.strip() != "")
            if isinstance(nonempty_mask, bool):
                nonempty_mask = pd.Series([True] * len(df), index=df.index)

            dups = df.loc[nonempty_mask].duplicated(subset=usable, keep=False)
            dup_idx = df.loc[nonempty_mask].index[dups].tolist()[:5000]
            for i in dup_idx:
                warnings.append(
                    {
                        "type": "duplicate",
                        "row": int(i),
                        "column": ",".join(usable),
                        "message": f"Possible duplicate row based on key: {usable}",
                    }
                )
                flagged.add(int(i))

    err_df = pd.DataFrame(errors)
    warn_df = pd.DataFrame(warnings)
    ok = len(err_df) == 0
    return ValidationResult(ok, err_df, warn_df, sorted(flagged))



# =============================================================================
# Extraction control helpers
# =============================================================================
def _first_existing_col(df: pd.DataFrame, candidates: List[str]) -> Optional[str]:
    for c in candidates:
        if c in df.columns:
            return c
    return None

def build_extraction_control_df(ui_df: pd.DataFrame) -> pd.DataFrame:
    """
    Fallback only. Real Extraction Control must come from the extractor,
    because the app does not have the original PDF path/raw text.
    """
    return pd.DataFrame(columns=[
        "Page number",
        "Raw line no",
        "Raw PDF text",
        "Matched extracted row",
        "Matched extracted text",
        "Match score",
        "Missing tokens",
        "Coverage status",
        "Control comment",
    ])

    
def get_extractor_control_from_df(df: pd.DataFrame) -> pd.DataFrame:
    if not isinstance(df, pd.DataFrame):
        return pd.DataFrame()

    for key in ("extraction_control", "Extraction control", "Extraction Control"):
        control = getattr(df, "attrs", {}).get(key)
        if isinstance(control, pd.DataFrame) and not control.empty:
            return control.copy()

    return build_extraction_control_df(df)


# =============================================================================
# Excel helpers (review sheet, parse, diff)
# =============================================================================
def make_review_sheet_df() -> pd.DataFrame:
    skipped = st.session_state.skipped_pages or []
    skipped_txt = ",".join(map(str, skipped)) if skipped else ""

    sec = st.session_state.extraction_seconds
    sec_txt = "" if sec is None else f"{sec:.2f}"
    rows = int(st.session_state.extraction_rows or 0)
    rate_txt = ""
    if sec is not None and sec > 0 and rows > 0:
        rate_txt = f"{(rows / sec):.2f} rows/s"

    return pd.DataFrame(
        [
            {"Field": "Site", "Value": st.session_state.site},
            {"Field": "Product", "Value": st.session_state.product_name},
            {"Field": "File name", "Value": st.session_state.batch_record_name or ""},
            {"Field": "Job ID", "Value": st.session_state.job_id},
            {"Field": "Reviewer", "Value": st.session_state.reviewer_name},
            {"Field": "Status", "Value": st.session_state.review_status},
            {"Field": "Reviewed at", "Value": st.session_state.reviewed_at},
            {"Field": "Total rows extracted", "Value": len(st.session_state.df) if st.session_state.df is not None else 0},
            {"Field": "Skipped pages", "Value": skipped_txt},
            {"Field": "Validation status", "Value": st.session_state.validation_status},
            {"Field": "Baseline hash", "Value": st.session_state.baseline_hash},
            {"Field": "Reviewer comment", "Value": st.session_state.review_comment},
            {"Field": "Extraction source", "Value": st.session_state.extraction_source},
            {"Field": "Extraction started at", "Value": st.session_state.extraction_started_at},
            {"Field": "Extraction ended at", "Value": st.session_state.extraction_ended_at},
            {"Field": "Extraction duration (s)", "Value": sec_txt},
            {"Field": "Extraction throughput", "Value": rate_txt},
        ]
    )


def parse_review_sheet(review_df: Optional[pd.DataFrame]) -> Dict[str, str]:
    if review_df is None or len(review_df) == 0:
        return {}
    cols = [c.lower().strip() for c in review_df.columns]
    if "field" not in cols or "value" not in cols:
        return {}
    field_col = review_df.columns[cols.index("field")]
    value_col = review_df.columns[cols.index("value")]
    out = {}
    for _, r in review_df.iterrows():
        k = str(r[field_col]).strip()
        v = "" if pd.isna(r[value_col]) else str(r[value_col])
        out[k] = v
    return out


def diff_summary(old: pd.DataFrame, new: pd.DataFrame, cols: List[str]) -> pd.DataFrame:
    out = []
    n = min(len(old), len(new))
    old2 = old.iloc[:n].reset_index(drop=True)
    new2 = new.iloc[:n].reset_index(drop=True)
    for col in cols:
        o = old2[col].astype(str) if col in old2.columns else pd.Series([""] * n)
        nn = new2[col].astype(str) if col in new2.columns else pd.Series([""] * n)
        out.append({"column": col, "changed_cells": int((o != nn).sum())})
    out.append({"column": "__row_count__", "changed_cells": int(abs(len(old) - len(new)))})
    return pd.DataFrame(out)


# =============================================================================
# Robust Excel sheet name handling
# =============================================================================
def _norm_sheet_name(s: str) -> str:
    return str(s).strip().lower()


def find_sheet_name(xl: pd.ExcelFile, desired: str) -> Optional[str]:
    """Exact match ignoring case and surrounding spaces."""
    desired_norm = _norm_sheet_name(desired)
    for s in xl.sheet_names:
        if _norm_sheet_name(s) == desired_norm:
            return s
    return None


def find_sheet_by_contains(xl: pd.ExcelFile, token: str) -> Optional[str]:
    """Fallback: first sheet name that contains token (case-insensitive)."""
    t = token.strip().lower()
    for s in xl.sheet_names:
        if t in _norm_sheet_name(s):
            return s
    return None


# =============================================================================
# Workflow columns normalization
# =============================================================================
REVIEW_OK_DEFAULTS = {
    REVIEW_OK_COL: False,
    REVIEWER_COL: "",
    REVIEWED_AT_COL: "",
    REVIEWER_NOTE_COL: "",
}

ALIASES_TO_CANONICAL = {
    "Needs review": "__OLD_PENDING__",
    "Pending review": "__OLD_PENDING__",
    "Pending Review": "__OLD_PENDING__",
    "Needs Review": "__OLD_PENDING__",
    "Reviewed and OK": REVIEW_OK_COL,
    "Reviewed OK": REVIEW_OK_COL,
    "Reviewed At": REVIEWED_AT_COL,
    "Review note": REVIEWER_NOTE_COL,
    "Reviewer Note": REVIEWER_NOTE_COL,
}


def coerce_review_ok_series(s: pd.Series) -> pd.Series:
    """
    Normalize Review OK values for the app checkbox.
    Important: bool("☐") is True in Python, so never use astype(bool)
    directly on this column.
    """
    def one(v):
        if pd.isna(v):
            return False
        if isinstance(v, bool):
            return v
        if isinstance(v, (int, float)):
            return bool(v)
        t = str(v).strip().lower()
        if t in {"", "false", "0", "no", "n", "unchecked", "☐", "□", "none", "nan"}:
            return False
        if t in {"true", "1", "yes", "y", "checked", "☑", "☒", "✓", "✔", "x"}:
            return True
        # Unknown text should default to False to avoid marking rows reviewed by mistake.
        return False

    return s.map(one).astype(bool)



def review_ok_to_excel_symbols(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()

    if REVIEW_OK_COL in out.columns:
        ok = coerce_review_ok_series(out[REVIEW_OK_COL])
        out[REVIEW_OK_COL] = ok.map(lambda x: "þ" if x else "¨")

    return out

def ensure_workflow_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.rename(columns={c: str(c).strip() for c in df.columns})
    rename_map = {c: ALIASES_TO_CANONICAL[c] for c in df.columns if c in ALIASES_TO_CANONICAL}
    if rename_map:
        df = df.rename(columns=rename_map)

    if "__OLD_PENDING__" in df.columns and REVIEW_OK_COL not in df.columns:
        old_pending = coerce_review_ok_series(df["__OLD_PENDING__"])
        df[REVIEW_OK_COL] = (~old_pending)
        df = df.drop(columns=["__OLD_PENDING__"])

    for col, default in REVIEW_OK_DEFAULTS.items():
        if col not in df.columns:
            df[col] = default

    df[REVIEW_OK_COL] = coerce_review_ok_series(df[REVIEW_OK_COL])
    return df


# =============================================================================
# Macro-enabled Excel export
# =============================================================================
# Path to the macro-enabled template (.xlsm). Created once, kept next to app.py.
# The template must contain three empty sheets ("Extracted", "Review", "Audit log")
# and the VBA Worksheet_BeforeDoubleClick on the "Extracted" sheet that
# toggles ☑/☐ and stamps the date in "Reviewed at".
MACRO_TEMPLATE_PATH = os.path.join(
    os.path.dirname(os.path.abspath(__file__)),
    "template.xlsm",
)

EXT_XLSM = "xlsm"
EXT_XLSX = "xlsx"
MIME_XLSM = "application/vnd.ms-excel.sheet.macroEnabled.12"
MIME_XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


def _macro_template_available() -> bool:
    return os.path.exists(MACRO_TEMPLATE_PATH)


def _clear_worksheet(ws) -> None:
    """Remove all rows but keep the sheet — preserves its VBA code module binding."""
    if ws.max_row and ws.max_row >= 1:
        ws.delete_rows(1, ws.max_row)


def _write_df_to_worksheet(ws, df: pd.DataFrame) -> None:
    """Write a DataFrame to an openpyxl worksheet starting at A1."""
    if df is None:
        return

    for j, col in enumerate(df.columns, start=1):
        ws.cell(row=1, column=j, value=str(col))

    if df.empty:
        return

    for i, row in enumerate(df.itertuples(index=False, name=None), start=2):
        for j, val in enumerate(row, start=1):
            if val is None or (isinstance(val, float) and pd.isna(val)):
                ws.cell(row=i, column=j, value=None)
            else:
                ws.cell(row=i, column=j, value=val)

    style_checklist_column(ws, REVIEW_OK_COL)


def build_excel_bytes_macro(
    extracted_df: pd.DataFrame,
    review_df: pd.DataFrame,
    audit_df: Optional[pd.DataFrame],
    include_audit: bool = True,
    control_df_override: Optional[pd.DataFrame] = None,
) -> Tuple[bytes, str, str]:
    """
    Build the export workbook.
    Returns (file_bytes, mime_type, extension).
    """
    export_df = review_ok_to_excel_symbols(extracted_df)

    if isinstance(control_df_override, pd.DataFrame):
        control_df = control_df_override
    else:
        control_df = (
            st.session_state.extraction_control_df
            if hasattr(st.session_state, "extraction_control_df")
            and isinstance(st.session_state.extraction_control_df, pd.DataFrame)
            else pd.DataFrame()
        )

    if control_df.empty:
        control_df = build_extraction_control_df(extracted_df)

    if _macro_template_available():
        try:
            wb = openpyxl.load_workbook(MACRO_TEMPLATE_PATH, keep_vba=True)

            for sheet_name, df_to_write in [
                ("Extracted", export_df),
                ("Review", review_df),
                ("Extraction control", control_df),
            ]:
                if sheet_name not in wb.sheetnames:
                    wb.create_sheet(sheet_name)

                ws = wb[sheet_name]
                _clear_worksheet(ws)
                _write_df_to_worksheet(ws, df_to_write)

            if include_audit:
                if "Audit log" not in wb.sheetnames:
                    wb.create_sheet("Audit log")

                ws = wb["Audit log"]
                _clear_worksheet(ws)
                _write_df_to_worksheet(
                    ws,
                    audit_df if audit_df is not None else pd.DataFrame()
                )

            out = io.BytesIO()
            wb.save(out)
            return out.getvalue(), MIME_XLSM, EXT_XLSM

        except Exception as e:
            audit("macro_template_failed_fallback_xlsx", f"{type(e).__name__}: {e}")

    # Fallback: plain .xlsx
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as w:
        export_df.to_excel(w, index=False, sheet_name="Extracted")
        review_df.to_excel(w, index=False, sheet_name="Review")
        control_df.to_excel(w, index=False, sheet_name="Extraction control")

        if include_audit:
            (audit_df if audit_df is not None else pd.DataFrame()).to_excel(
                w,
                index=False,
                sheet_name="Audit log"
            )

    return out.getvalue(), MIME_XLSX, EXT_XLSX


def df_to_excel_bytes(df: pd.DataFrame, include_audit: bool = True) -> Tuple[bytes, str, str]:
    """
    Thin wrapper used by the review-pack download in Step 2.
    Returns (file_bytes, mime_type, extension).
    """
    propagate_reviewer_to_df(fill_all_rows=False)
    review_df = make_review_sheet_df()
    audit_df = pd.DataFrame(st.session_state.audit_log) if include_audit else None
    return build_excel_bytes_macro(df, review_df, audit_df, include_audit)


# =============================================================================
# Extraction plumbing
# =============================================================================
def fp(b: bytes) -> str:
    return f"{len(b)}:{hash(b[:2048])}"


def extract_by_site(site: str, file_bytes: bytes, file_name: str) -> pd.DataFrame:

    extractor = get_extractor(site)
    df = extractor(file_bytes, file_name)

    # Add missing review/workflow columns
    df = ensure_workflow_columns(df)

    # Fill UI metadata for Arklow / Bolbec / Toledo
    df = apply_ui_metadata_to_df(site, df, file_name)

    # Build strong raw-text control BEFORE schema trimming
    try:
        import tempfile

        existing_control = getattr(df, "attrs", {}).get("extraction_control")

        if not isinstance(existing_control, pd.DataFrame) or existing_control.empty:

            with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
                tmp.write(file_bytes)
                tmp_pdf_path = tmp.name

            from extraction_control import build_strong_raw_order_control

            control_df = build_strong_raw_order_control(tmp_pdf_path, df)
            df.attrs["extraction_control"] = control_df

            try:
                os.remove(tmp_pdf_path)
            except Exception:
                pass

    except Exception as e:
        audit("raw_text_control_failed", f"{site}: {type(e).__name__}: {e}")

    # Apply site schema columns
    schema = get_active_schema(site)
    cols = schema.get("columns") or list(df.columns)

    if schema.get("columns"):
        missing = [c for c in cols if c not in df.columns]
        if missing:
            raise ValueError(f"Extractor for {site} returned missing columns: {missing}")
        df = df[cols].copy()

    return df

def maybe_auto_extract():
    ss = st.session_state
    if ss.batch_record_bytes is None:
        return

    new_fp = fp(ss.batch_record_bytes)

    if ss.batch_record_fp == new_fp and isinstance(ss.df, pd.DataFrame):
        ss.df = apply_ui_metadata_to_df(ss.site, ss.df, ss.batch_record_name or "record")
        return

    ss.extraction_source = "pdf_word"
    ss.extraction_started_at = time.strftime("%Y-%m-%d %H:%M:%S")
    t0 = time.perf_counter()

    schema = get_active_schema(ss.site)
    df = extract_by_site(ss.site, ss.batch_record_bytes, ss.batch_record_name or "record")

    t1 = time.perf_counter()
    ss.extraction_ended_at = time.strftime("%Y-%m-%d %H:%M:%S")
    ss.extraction_seconds = float(t1 - t0)
    ss.extraction_rows = int(len(df))

    ss.df = df
    ss.extraction_control_df = get_extractor_control_from_df(df)
    ss.active_schema = schema
    ss.active_columns = derive_columns_from_schema_or_df(ss.df, ss.active_schema)

    ss.validation = validate_df(ss.df, ss.active_schema)
    ss.user_validated = False
    ss.validation_status = "Not validated yet"

    ss.batch_record_fp = new_fp

    ss.baseline_hash = df_content_hash(ss.df)
    ss.review_status = "Draft"
    ss.reviewed_at = ""
    ss.review_comment = ""

    audit(
        "auto_extracted",
        f"site={ss.site}; rows={len(ss.df)}; file={ss.batch_record_name}; seconds={ss.extraction_seconds:.2f}",
    )


def load_review_pack_excel(file_bytes: bytes, expected_cols: List[str]) -> Tuple[pd.DataFrame, Optional[pd.DataFrame]]:
    bio = io.BytesIO(file_bytes)
    xl = pd.ExcelFile(bio, engine="openpyxl")

    extracted_sheet = find_sheet_name(xl, "Extracted") or find_sheet_by_contains(xl, "extract")
    if not extracted_sheet:
        raise ValueError(
            "Missing sheet 'Extracted' in uploaded Excel. "
            f"Found sheets: {xl.sheet_names}. "
            "Please upload the review pack exported from this tool."
        )

    review_sheet = find_sheet_name(xl, "Review") or find_sheet_by_contains(xl, "review")

    extracted = pd.read_excel(xl, extracted_sheet, engine="openpyxl")
    review = pd.read_excel(xl, review_sheet, engine="openpyxl") if review_sheet else None

    extracted = ensure_workflow_columns(extracted)

    missing = [c for c in expected_cols if c not in extracted.columns]
    if missing:
        raise ValueError(f"Uploaded Excel missing required columns: {missing}")

    return extracted[expected_cols].copy(), review


def resume_from_review_pack(file_bytes: bytes, file_name: str):
    bio = io.BytesIO(file_bytes)
    xl = pd.ExcelFile(bio, engine="openpyxl")

    extracted_sheet = find_sheet_name(xl, "Extracted") or find_sheet_by_contains(xl, "extract")
    if not extracted_sheet:
        raise ValueError(
            "Missing sheet 'Extracted' in uploaded Excel. "
            f"Found sheets: {xl.sheet_names}. "
            "Please upload the review pack exported from this tool."
        )

    review_sheet = find_sheet_name(xl, "Review") or find_sheet_by_contains(xl, "review")
    control_sheet = (
        find_sheet_name(xl, "Extraction control")
        or find_sheet_name(xl, "Extraction Control")
        or find_sheet_by_contains(xl, "control")
    )

    extracted = pd.read_excel(xl, extracted_sheet, engine="openpyxl")
    review_df = pd.read_excel(xl, review_sheet, engine="openpyxl") if review_sheet else None
    control_df = pd.read_excel(xl, control_sheet, engine="openpyxl") if control_sheet else pd.DataFrame()
    meta = parse_review_sheet(review_df)

    site = meta.get("Site", "") or st.session_state.site
    if site not in SITES:
        site = st.session_state.site

    st.session_state.site = site
    st.session_state.product_name = meta.get("Product", "") or st.session_state.product_name
    st.session_state.batch_record_name = meta.get("File name", "") or file_name

    st.session_state.active_schema = get_active_schema(site)
    expected_cols = derive_columns_from_schema_or_df(None, st.session_state.active_schema)

    extracted = ensure_workflow_columns(extracted)

    extracted = apply_ui_metadata_to_df(site, extracted, st.session_state.batch_record_name or file_name)

    missing = [c for c in expected_cols if c not in extracted.columns]
    if missing:
        raise ValueError(f"Uploaded Excel missing required columns for site '{site}': {missing}")

    st.session_state.df = extracted[expected_cols].copy()
    st.session_state.extraction_control_df = (
        control_df if isinstance(control_df, pd.DataFrame) and not control_df.empty
        else build_extraction_control_df(st.session_state.df)
    )
    st.session_state.active_columns = expected_cols

    st.session_state.validation = validate_df(st.session_state.df, st.session_state.active_schema)
    st.session_state.user_validated = False
    st.session_state.validation_status = "Not validated yet"

    st.session_state.baseline_hash = df_content_hash(st.session_state.df)

    st.session_state.review_status = "In review"
    st.session_state.reviewed_at = meta.get("Reviewed at", "") or ""
    st.session_state.review_comment = meta.get("Reviewer comment", "") or ""
    st.session_state.reviewer_name = meta.get("Reviewer", "") or st.session_state.reviewer_name

    st.session_state.extraction_source = meta.get("Extraction source", "") or "review_pack"
    st.session_state.extraction_started_at = meta.get("Extraction started at", "") or ""
    st.session_state.extraction_ended_at = meta.get("Extraction ended at", "") or ""
    try:
        sec = meta.get("Extraction duration (s)", "")
        st.session_state.extraction_seconds = float(sec) if str(sec).strip() else None
    except Exception:
        st.session_state.extraction_seconds = None

    st.session_state.extraction_rows = int(len(st.session_state.df))

    if needs_product():
        set_blocking_messages(["• Please enter **Product name** first (Step 1)."])
        st.session_state.view_step = "Setup"
        audit("resumed_from_review_pack_missing_product", f"file={file_name}; site={site}; rows={len(st.session_state.df)}")
    else:
        set_blocking_messages([])
        st.session_state.view_step = "Review"
        audit("resumed_from_review_pack", f"file={file_name}; site={site}; rows={len(st.session_state.df)}")


# =============================================================================
# Top bar
# =============================================================================
top_left, top_center, top_right = st.columns([1, 8, 1])

with top_left:
    st.markdown('<div class="kebab-btn">', unsafe_allow_html=True)
    if st.button("⋮", help="Menu"):
        st.session_state.show_menu = not st.session_state.show_menu
    st.markdown("</div>", unsafe_allow_html=True)

with top_center:
    st.markdown(
        """
        <div style="text-align:center; margin-top: 6px;">
          <div style="
            font-size: 46px;
            font-weight: 1000;
            letter-spacing: -1px;
            color: #ffffff;
            text-shadow: 0 3px 10px rgba(0,0,0,0.55);
            text-transform: uppercase;
          ">
            Auto BR Extractor Tool
          </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

with top_right:
    st.write("")

if st.session_state.show_menu:
    with st.expander("Menu", expanded=True):

        st.markdown("### Simplification support")
        st.write(
            """
This tool is designed to support the **Batch Record Simplification** exercise by extracting structured data
from **empty Batch Record templates**, with the output aligned to the toolbox template structure.
            """
        )

        st.markdown("---")

        st.markdown("### About this tool")
        st.write(
            """
- Start either from Batch Record (PDF) OR from Review Pack (Excel)
- Review & correct
- Validation status only changes when you click "Validate table" or "Approve & lock"
- Approve + Export final Excel
- Downloaded Excel files are macro-enabled (.xlsm) when template.xlsm is present:
  double-click the "Reviewed & OK" cell to toggle ☑ and stamp the date automatically.
            """
        )

        st.markdown("---")
        st.button("🔄 Reset session", on_click=reset_all)
        st.write()

# =============================================================================
# Steps nav (ALWAYS ENABLED)
# =============================================================================
b1, b2, b3, b4, b5 = st.columns([2, 1, 1, 1, 2])
with b2:
    go_setup = st.button("Step 1 · Setup")
with b3:
    go_review = st.button("Step 2 · Review")
with b4:
    go_export = st.button("Step 3 · Export Excel")

if go_setup:
    set_blocking_messages([])
    st.session_state.view_step = "Setup"
elif go_review:
    go_or_warn("Review")
elif go_export:
    go_or_warn("Export Excel")

enforce_step_guard_soft()


# =============================================================================
# Summary cards
# =============================================================================
def render_summary_cards():
    df = st.session_state.df
    schema = st.session_state.active_schema or {}

    rows = f"{len(df):,}" if isinstance(df, pd.DataFrame) else "-"
    skipped = st.session_state.skipped_pages or []
    skipped_txt = ", ".join(map(str, skipped)) if skipped else "-"

    page_count_txt = "-"
    page_col = (schema.get("page_col") or "").strip()
    if isinstance(df, pd.DataFrame) and page_col and page_col in df.columns:
        pages = pd.to_numeric(df[page_col], errors="coerce").dropna()
        if len(pages) > 0:
            page_count_txt = str(int(pages.round().astype(int).nunique()))

    sec = st.session_state.extraction_seconds
    source = st.session_state.extraction_source or "-"
    timing_txt = "-"
    if sec is not None:
        timing_txt = f"{sec:.2f}s"
        if st.session_state.extraction_rows and sec > 0:
            timing_txt += f" · {(st.session_state.extraction_rows / sec):.2f} rows/s"

    k1, k2, k3, k4, k5, k6, k7 = st.columns(7)
    with k1:
        st.markdown(
            f'<div class="metric-card"><div class="metric-label">File name</div><div class="metric-value">{st.session_state.batch_record_name or "-"}</div></div>',
            unsafe_allow_html=True,
        )
    with k2:
        st.markdown(
            f'<div class="metric-card"><div class="metric-label">Date / user</div><div class="metric-value">{time.strftime("%Y-%m-%d")} · {st.session_state.reviewer_name or "-"}</div></div>',
            unsafe_allow_html=True,
        )
    with k3:
        st.markdown(
            f'<div class="metric-card"><div class="metric-label">Total rows</div><div class="metric-value">{rows}</div></div>',
            unsafe_allow_html=True,
        )
    with k4:
        st.markdown(
            f'<div class="metric-card">'
            f'  <div class="metric-label">Product name</div>'
            f'  <div class="metric-value">{st.session_state.product_name or "-"}</div>'
            f'</div>',
            unsafe_allow_html=True
        )
    with k5:
        st.markdown(
            f'<div class="metric-card"><div class="metric-label">Number of pages</div><div class="metric-value">{page_count_txt}</div></div>',
            unsafe_allow_html=True,
        )
    with k6:
        st.markdown(
            f'<div class="metric-card"><div class="metric-label">Skipped pages</div><div class="metric-value">{skipped_txt}</div></div>',
            unsafe_allow_html=True,
        )
    with k7:
        st.markdown(
            f'<div class="metric-card"><div class="metric-label">Extraction</div><div class="metric-value">{source} · {timing_txt}</div></div>',
            unsafe_allow_html=True,
        )

    st.markdown('<div class="hr"></div>', unsafe_allow_html=True)
    st.markdown(
        f'<div class="metric-card"><div class="metric-label">Validation</div>'
        f'<div class="metric-value">{st.session_state.validation_status} · Status: {st.session_state.review_status}</div></div>',
        unsafe_allow_html=True,
    )
    st.markdown('<div class="hr"></div>', unsafe_allow_html=True)


# =============================================================================
# Step 2 helper: reviewer propagation
# =============================================================================
def propagate_reviewer_to_df(fill_all_rows: bool = False):
    df = st.session_state.df
    if not isinstance(df, pd.DataFrame):
        return

    schema = st.session_state.active_schema or {}
    ok_col = (schema.get("review_ok_col") or REVIEW_OK_COL).strip()
    reviewer_col = (schema.get("reviewer_col") or REVIEWER_COL).strip()
    reviewed_at_col = (schema.get("reviewed_at_col") or REVIEWED_AT_COL).strip()

    who = (st.session_state.reviewer_name or "").strip()
    if not who:
        return
    if reviewer_col not in df.columns:
        return

    if fill_all_rows:
        target_idx = df.index
    else:
        empty_mask = df[reviewer_col].isna() | (df[reviewer_col].astype(str).str.strip() == "")
        target_idx = df.index[empty_mask]

    if len(target_idx) > 0:
        df.loc[target_idx, reviewer_col] = who

    if ok_col in df.columns and reviewed_at_col in df.columns:
        ok_mask = coerce_review_ok_series(df[ok_col])
        missing_dt = df[reviewed_at_col].isna() | (df[reviewed_at_col].astype(str).str.strip() == "")
        dt_idx = df.index[ok_mask & missing_dt]
        if len(dt_idx) > 0:
            df.loc[dt_idx, reviewed_at_col] = time.strftime("%Y-%m-%d %H:%M:%S")

    st.session_state.df = df


# =============================================================================
# Step 1: Setup
# =============================================================================
def render_setup_step():
    st.markdown('<div class="panel">', unsafe_allow_html=True)
    step_title("Setup", "Step 1: Start new OR resume from Excel")

    if st.session_state.blocking_msgs:
        st.warning("You can’t continue yet:\n\n" + "\n".join(st.session_state.blocking_msgs))

    c1, c2 = st.columns([1.2, 1.8])
    with c1:
        new_site = st.selectbox("Select site", SITES, index=SITES.index(st.session_state.site))
    with c2:
        new_product = st.text_input(
            "Product name (required)",
            value=st.session_state.product_name,
            placeholder="Type product name here…",
        )

    if new_product != st.session_state.product_name:
        st.session_state.product_name = new_product
        if isinstance(st.session_state.df, pd.DataFrame):
            st.session_state.df = apply_ui_metadata_to_df(
                st.session_state.site,
                st.session_state.df,
                st.session_state.batch_record_name or "record",
            )

    if new_site != st.session_state.site:
        st.session_state.site = new_site
        st.session_state.active_schema = get_active_schema(new_site)
        st.session_state.active_columns = derive_columns_from_schema_or_df(st.session_state.df, st.session_state.active_schema)
        st.session_state.batch_record_fp = None

        if isinstance(st.session_state.df, pd.DataFrame):
            st.session_state.df = apply_ui_metadata_to_df(
                st.session_state.site,
                st.session_state.df,
                st.session_state.batch_record_name or "record",
            )

    tab_new, tab_resume = st.tabs(["New extraction (PDF)", "Resume from Review Pack (Excel)"])

    with tab_new:
        st.markdown("#### Upload Batch Record (single file)")
        st.markdown(
            """
            <div style="
                background: rgba(10, 18, 35, 0.75);
                border: 1px solid rgba(255,255,255,0.25);
                border-radius: 16px;
                padding: 18px;
                box-shadow: 0 10px 28px rgba(0,0,0,0.35);
                backdrop-filter: blur(10px);
                margin-bottom: 15px;
            ">
                <div style="font-weight:900; font-size:18px; margin-bottom:10px;">
                    Important notice
                </div>
                <ul style="margin:0; padding-left:20px; line-height:1.6;">
                    <li>This tool works only with <b>EMPTY Batch Record templates</b>.</li>
                    <li>Do NOT upload filled, handwritten, or completed batch records.</li>
                    <li>Scanned PDFs (image-based files) can alter extraction quality and structure.</li>
                    <li>Please upload the original digital empty template in PDF format.</li>
                </ul>
            </div>
            """,
            unsafe_allow_html=True,
        )
        uploaded = st.file_uploader(
            label="Upload Batch Record (PDF)",
            type=["pdf"],
            key="uploader_batch_record",
            help="Extraction starts automatically after upload.",
        )
        if uploaded is not None:
            st.session_state.batch_record_bytes = uploaded.read()
            st.session_state.batch_record_name = uploaded.name
            audit("batch_record_uploaded", uploaded.name)
            with st.spinner("Auto-extracting..."):
                time.sleep(0.05)
                maybe_auto_extract()
            st.success("Loaded and extracted ✅ (You can now go to Step 2 · Review)")
            set_blocking_messages([])

    with tab_resume:
        st.markdown("#### Upload Review Pack (Excel) to resume review (even next day)")
        uploaded_xlsx = st.file_uploader(
            label="Upload Review Pack (Excel)",
            type=["xlsx", "xlsm"],
            key="uploader_review_pack",
            help="Upload the Excel pack previously downloaded from this tool (.xlsx or .xlsm).",
            label_visibility="collapsed",
        )
        if uploaded_xlsx is not None:
            try:
                st.session_state.extraction_source = "review_pack"
                resume_from_review_pack(uploaded_xlsx.read(), uploaded_xlsx.name)
                if needs_product():
                    st.warning("Review pack loaded ✅ Now please enter Product name (required) before going to Step 2.")
                else:
                    st.success("Review pack loaded ✅ You are now in Step 2 · Review.")
                st.rerun()
            except Exception as e:
                st.error(f"Could not resume from review pack: {e}")
                st.code(traceback.format_exc())

    st.markdown('<div class="hr"></div>', unsafe_allow_html=True)
    next_disabled = needs_product() or needs_data()

    col_next = st.columns([1])[0]
    with col_next:
        if st.button("➡ Next: Go to Step 2 · Review", disabled=next_disabled, use_container_width=True):
            set_blocking_messages([])
            st.session_state.view_step = "Review"
            st.rerun()

        if next_disabled:
            missing = []
            if needs_product():
                missing.append("Product name")
            if needs_data():
                missing.append("Upload Batch Record / Review Pack")
            st.caption("Complete required fields: " + ", ".join(missing))

    st.markdown("</div>", unsafe_allow_html=True)


# =============================================================================
# Step 2: Review
# =============================================================================
def render_review_metadata_panel():
    st.subheader("Reviewer / status")

    cA, cB, cC = st.columns([1.3, 1.0, 1.0])

    with cA:
        new_reviewer = st.text_input(
            "Reviewer name",
            value=st.session_state.reviewer_name,
            placeholder="Required for approval/export/download",
        )

    if new_reviewer != st.session_state.reviewer_name:
        st.session_state.reviewer_name = new_reviewer
        propagate_reviewer_to_df(fill_all_rows=False)
        mark_dirty("reviewer_name_changed")
        audit("reviewer_name_changed", f"reviewer={new_reviewer}")

    with cB:
        st.session_state.review_status = st.selectbox(
            "Review status",
            REVIEW_STATUS_CHOICES,
            index=REVIEW_STATUS_CHOICES.index(st.session_state.review_status),
        )

    with cC:
        st.caption(f"Reviewed at: {st.session_state.reviewed_at or '-'}")

    st.session_state.review_comment = st.text_area(
        "Reviewer global comment",
        value=st.session_state.review_comment,
        placeholder="Optional comment...",
        height=80,
    )

    c1, c2, c3 = st.columns([1.1, 1.1, 2.0])

    with c1:
        if st.button("✅ Approve & lock"):
            if needs_product():
                st.error("Product name is required to approve. Please fill it in Step 1.")
                return

            if not st.session_state.reviewer_name.strip():
                st.error("Reviewer name is required to approve.")
                return

            st.session_state.validation = validate_df(
                st.session_state.df,
                st.session_state.active_schema,
            )
            st.session_state.user_validated = True

            err_n = len(st.session_state.validation.errors)
            warn_n = len(st.session_state.validation.warnings)

            if err_n > 0:
                st.session_state.validation_status = f"Validated with notes ⚠️ ({err_n} items)"
                st.error("Approval blocked: some required fields are missing.")
                with st.expander("Show details"):
                    st.dataframe(
                        st.session_state.validation.errors,
                        use_container_width=True,
                        hide_index=True,
                    )
                audit("validated_on_approve_blocked", f"errors={err_n}; warnings={warn_n}")
                return

            st.session_state.validation_status = "Validated ✅"

            df = st.session_state.df
            schema = st.session_state.active_schema or {}
            ok_col = (schema.get("review_ok_col") or REVIEW_OK_COL).strip()

            not_reviewed = (
                int((~coerce_review_ok_series(df[ok_col])).sum())
                if df is not None and ok_col in df.columns
                else 0
            )

            if not_reviewed > 0:
                st.error(f"Approval blocked: {not_reviewed} rows are not reviewed yet.")
                return

            propagate_reviewer_to_df(fill_all_rows=False)

            st.session_state.review_status = "Approved"
            st.session_state.reviewed_at = time.strftime("%Y-%m-%d %H:%M:%S")
            st.session_state.baseline_hash = df_content_hash(st.session_state.df)

            audit(
                "approved",
                f"reviewer={st.session_state.reviewer_name}; hash={st.session_state.baseline_hash}",
            )

            st.success("Approved and locked ✅")
            st.rerun()

    with c2:
        if st.button("🔓 Unlock (set Draft)"):
            st.session_state.review_status = "Draft"
            st.session_state.reviewed_at = ""
            audit("unlocked", "status set to Draft")
            st.rerun()

    with c3:
        st.caption(
            "Tip: Check 'Reviewed & OK' when a row is verified. Reviewer and timestamp will be auto-filled."
        )


def render_review_pack_panel():
    st.subheader("Offline review (recommended for 1000+ rows)")

    cols = st.session_state.active_columns or (
        list(st.session_state.df.columns) if st.session_state.df is not None else []
    )

    c1, c2 = st.columns([1.2, 1.0])

    with c1:
        if st.session_state.df is not None:

            reviewer_ok = bool((st.session_state.reviewer_name or "").strip())

            if reviewer_ok:
                template_hint = (
                    "Macro-enabled (.xlsm): double-click 'Reviewed & OK' to toggle ☑ and stamp the date."
                    if _macro_template_available()
                    else "Note: template.xlsm not found — file will be plain .xlsx (no macros)."
                )

                st.markdown(
                    f"""
                    <div style="
                        background: rgba(0, 180, 120, 0.15);
                        border: 1px solid rgba(0,255,170,0.4);
                        border-radius: 12px;
                        padding: 10px;
                        font-weight: 900;
                        color: white;
                    ">
                        ✅ Reviewer name entered. Review pack is ready to download.<br>
                        <span style="font-weight:700; opacity:0.95;">{template_hint}</span>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )

                propagate_reviewer_to_df(fill_all_rows=False)

                pack_bytes, pack_mime, pack_ext = df_to_excel_bytes(
                    st.session_state.df, include_audit=True
                )
                pack_name = f"{st.session_state.site}_review_pack_{time.strftime('%Y%m%d_%H%M')}.{pack_ext}"

                st.download_button(
                    label="📤 Download review pack (Excel)",
                    data=pack_bytes,
                    file_name=pack_name,
                    mime=pack_mime,
                    use_container_width=True,
                )

            else:
                st.markdown(
                    """
                    <div style="
                        background: rgba(255, 255, 255, 0.08);
                        border: 1px solid rgba(255,255,255,0.3);
                        border-radius: 12px;
                        padding: 10px;
                        font-weight: 900;
                        color: white;
                    ">
                        ℹ️ Please enter reviewer name above to enable download.
                    </div>
                    """,
                    unsafe_allow_html=True,
                )

                st.button(
                    "📤 Download review pack (Excel)",
                    disabled=True,
                    use_container_width=True,
                )
    with c2:
        uploaded_xlsx = st.file_uploader(
            "📥 Upload reviewed Excel (review pack)",
            type=["xlsx", "xlsm"],
            key="upload_review_pack_inside_review"
        )

    if uploaded_xlsx is not None:
        try:
            new_df, _review_df = load_review_pack_excel(uploaded_xlsx.read(), expected_cols=cols)

            old_df = st.session_state.df.copy() if st.session_state.df is not None else None
            st.session_state.df = new_df
            mark_dirty("review_pack_upload")

            st.session_state.extraction_source = "review_pack"
            st.session_state.extraction_rows = int(len(new_df))

            st.session_state.validation = validate_df(
                st.session_state.df,
                st.session_state.active_schema
            )
            st.session_state.review_status = "In review"
            audit("review_pack_uploaded", f"name={uploaded_xlsx.name}; rows={len(new_df)}")

            st.success("Reviewed Excel uploaded ✅ Now click 'Validate table' when ready.")

            if old_df is not None:
                st.caption("Change summary (counts of changed cells per column):")
                st.dataframe(
                    diff_summary(old_df, new_df, cols),
                    use_container_width=True,
                    hide_index=True
                )

        except Exception as e:
            st.error(f"Could not import reviewed Excel: {e}")
            st.code(traceback.format_exc())


def render_review_step():
    st.markdown('<div class="panel">', unsafe_allow_html=True)
    step_title("Review", "Step 2: Review and correct")

    msgs = []
    if needs_product():
        msgs.append("• Please enter **Product name** first (Step 1).")
    if needs_data():
        msgs.append("• Please **upload Batch Record (PDF/Word)** OR **upload Review Pack (Excel)** first (Step 1).")
    if msgs:
        st.warning("You can’t continue yet:\n\n" + "\n".join(msgs))
        st.markdown("</div>", unsafe_allow_html=True)
        return

    render_summary_cards()
    render_review_metadata_panel()
    st.markdown('<div class="hr"></div>', unsafe_allow_html=True)
    render_review_pack_panel()
    st.markdown('<div class="hr"></div>', unsafe_allow_html=True)

    df = st.session_state.df
    schema = st.session_state.active_schema or {}
    cols = st.session_state.active_columns or list(df.columns)
    search_cols = derive_search_cols(schema, cols)

    ok_col = (schema.get("review_ok_col") or REVIEW_OK_COL).strip()
    reviewer_col = (schema.get("reviewer_col") or REVIEWER_COL).strip()
    reviewed_at_col = (schema.get("reviewed_at_col") or REVIEWED_AT_COL).strip()

    # Force Review OK to real boolean immediately before rendering.
    # This is required because values may arrive from Excel/extractor as strings
    # like "false", "☐", or "☑". Streamlit only renders CheckboxColumn
    # when the column is boolean, not object/string.
    if ok_col in df.columns:
        df[ok_col] = coerce_review_ok_series(df[ok_col])
        st.session_state.df = df

    is_locked = (st.session_state.review_status == "Approved")

    f1, f2 = st.columns([1.5, 2.2])
    with f1:
        st.session_state.view_mode = st.selectbox(
            "View mode",
            ["All rows", "Not reviewed only"],
            index=1,
            disabled=is_locked,
        )
    with f2:
        st.session_state.search = st.text_input(
            "Search",
            value=st.session_state.search,
            placeholder="Search any column…",
            disabled=is_locked
        )

    view = df.copy()

    if st.session_state.view_mode == "Not reviewed only":
        if ok_col in view.columns:
            view[ok_col] = coerce_review_ok_series(view[ok_col])
            view = view.loc[~view[ok_col]]
        else:
            st.warning(f"Column '{ok_col}' not found for this site.")

    q = (st.session_state.search or "").strip().lower()
    if q:
        mask = False
        for c in search_cols:
            if c in view.columns:
                mask = mask | view[c].astype(str).str.lower().str.contains(q, na=False)
        view = view[mask]

    view = view[cols].copy()
    if ok_col in view.columns:
        view[ok_col] = coerce_review_ok_series(view[ok_col])

    total = len(view)
    st.caption(f"Showing all rows: {total:,}")

    if is_locked:
        st.info("Locked because status is Approved. Click 'Unlock (set Draft)' to edit.")
    else:
        st.caption("Check 'Reviewed & OK' when a row is reviewed. After edits, click 'Validate table' (QA action).")

    page_view = view.copy()
    if ok_col in page_view.columns:
        page_view[ok_col] = coerce_review_ok_series(page_view[ok_col])
    before_page = df.loc[page_view.index].copy()
    if ok_col in before_page.columns:
        before_page[ok_col] = coerce_review_ok_series(before_page[ok_col])

    TABLE_HEIGHT = 900

    # Use Streamlit data_editor for the review table because it reliably renders
    # boolean columns as real checkboxes. AG Grid can display booleans as text
    # unless custom checkbox renderers are configured.
    if False and HAS_AGGRID:
        gb = GridOptionsBuilder.from_dataframe(page_view)
        gb.configure_default_column(editable=not is_locked, filter=True, sortable=True, resizable=True)
        gb.configure_pagination(enabled=False)
        grid_options = gb.build()
        grid_options["suppressCsvExport"] = True
        grid_options["suppressExcelExport"] = True

        grid = AgGrid(
            page_view,
            gridOptions=grid_options,
            update_mode=GridUpdateMode.VALUE_CHANGED,
            data_return_mode=DataReturnMode.FILTERED_AND_SORTED,
            fit_columns_on_grid_load=True,
            height=TABLE_HEIGHT,
            allow_unsafe_jscode=False,
        )
        edited_page = pd.DataFrame(grid["data"])
    else:
        edited_page = st.data_editor(
            page_view,
            use_container_width=True,
            hide_index=False,
            num_rows="fixed",
            height=TABLE_HEIGHT,
            disabled=is_locked,
            column_config={
                ok_col: st.column_config.CheckboxColumn(
                    "Reviewed & OK",
                    help="Check = reviewed and accepted.",
                )
            },
        )

    if not is_locked:
        try:
            edited_page.index = page_view.index
            for col in edited_page.columns:
                df.loc[edited_page.index, col] = edited_page[col]

            if ok_col in df.columns:
                df[ok_col] = coerce_review_ok_series(df[ok_col])
                was_ok = coerce_review_ok_series(before_page[ok_col])
                is_ok = coerce_review_ok_series(df.loc[edited_page.index, ok_col])

                newly_ok_idx = edited_page.index[(was_ok == False) & (is_ok == True)].tolist()
                if newly_ok_idx:
                    who = (st.session_state.reviewer_name or "").strip()
                    now_ts = time.strftime("%Y-%m-%d %H:%M:%S")

                    if reviewer_col in df.columns and who:
                        empty_reviewer = df.loc[newly_ok_idx, reviewer_col].astype(str).str.strip() == ""
                        df.loc[df.loc[newly_ok_idx].index[empty_reviewer], reviewer_col] = who

                    if reviewed_at_col in df.columns:
                        empty_dt = df.loc[newly_ok_idx, reviewed_at_col].astype(str).str.strip() == ""
                        df.loc[df.loc[newly_ok_idx].index[empty_dt], reviewed_at_col] = now_ts

            st.session_state.df = df
            mark_dirty("live_edit")
        except Exception as e:
            st.error(f"Could not apply edits: {e}")
            st.code(traceback.format_exc())

    if st.button("✅ Validate table", disabled=is_locked):
        st.session_state.validation = validate_df(st.session_state.df, schema)
        st.session_state.baseline_hash = df_content_hash(st.session_state.df)
        st.session_state.user_validated = True

        err_n = len(st.session_state.validation.errors)
        if err_n == 0:
            st.session_state.validation_status = "Validated ✅"
            st.success("Validation saved ✅ (no blocking issues)")
        else:
            st.session_state.validation_status = f"Validated with notes ⚠️ ({err_n} items)"
            st.info(f"Validation saved ✅ ({err_n} items flagged)")

        audit("user_validated", f"errors={err_n}; warnings={len(st.session_state.validation.warnings)}")

    st.markdown("</div>", unsafe_allow_html=True)


# =============================================================================
# Step 3: Export (no freeze; cache bytes)
# =============================================================================
@st.cache_data(show_spinner=False)
def _build_excel_bytes_cached(
    df_csv: str,
    include_audit: bool,
    audit_csv: str,
    review_csv: str,
    control_csv: str,
    template_marker: str,
) -> Tuple[bytes, str, str]:
    """
    Returns (file_bytes, mime_type, extension).

    template_marker is included only to invalidate the cache when the
    template.xlsm file appears/disappears or is replaced — its value is
    derived from the template path + its mtime.
    """
    df = pd.read_csv(io.StringIO(df_csv))
    review_df = pd.read_csv(io.StringIO(review_csv)) if review_csv else pd.DataFrame()
    audit_df = pd.read_csv(io.StringIO(audit_csv)) if (include_audit and audit_csv) else None
    control_df = pd.read_csv(io.StringIO(control_csv)) if control_csv else build_extraction_control_df(df)
    return build_excel_bytes_macro(df, review_df, audit_df, include_audit, control_df_override=control_df)


def _template_cache_marker() -> str:
    """Cheap fingerprint of template.xlsm so the cached export refreshes if it changes."""
    if not _macro_template_available():
        return "no-template"
    try:
        st_info = os.stat(MACRO_TEMPLATE_PATH)
        return f"{MACRO_TEMPLATE_PATH}:{int(st_info.st_mtime)}:{st_info.st_size}"
    except Exception:
        return "template-stat-failed"


def render_export_step():
    st.markdown('<div class="panel">', unsafe_allow_html=True)
    step_title("Export Excel", "Step 3: Export Excel file")

    msgs = []
    if needs_product():
        msgs.append("• Please enter **Product name** first (Step 1).")
    if needs_data():
        msgs.append("• Please **upload Batch Record (PDF/Word)** OR **upload Review Pack (Excel)** first (Step 1).")
    if msgs:
        st.warning("You can’t continue yet:\n\n" + "\n".join(msgs))
        st.markdown("</div>", unsafe_allow_html=True)
        return

    if st.session_state.review_status != "Approved":
        st.error("Export blocked: review status must be Approved.")
        st.info("Go to Step 2 · Review and click 'Approve & lock'.")
        st.markdown("</div>", unsafe_allow_html=True)
        return

    if not st.session_state.reviewer_name.strip():
        st.error("Export blocked: reviewer name is required.")
        st.markdown("</div>", unsafe_allow_html=True)
        return

    propagate_reviewer_to_df(fill_all_rows=False)

    df_csv = st.session_state.df.to_csv(index=False)
    review_df = make_review_sheet_df()
    review_csv = review_df.to_csv(index=False)
    audit_df = pd.DataFrame(st.session_state.audit_log)
    audit_csv = audit_df.to_csv(index=False)

    control_df = (
        st.session_state.extraction_control_df
        if isinstance(st.session_state.extraction_control_df, pd.DataFrame)
        and not st.session_state.extraction_control_df.empty
        else build_extraction_control_df(st.session_state.df)
    )
    control_csv = control_df.to_csv(index=False)

    with st.spinner("Preparing Excel…"):
        excel_bytes, excel_mime, excel_ext = _build_excel_bytes_cached(
            df_csv, True, audit_csv, review_csv, control_csv, _template_cache_marker()
        )

    filename = f"{st.session_state.site}_FINAL_{time.strftime('%Y%m%d_%H%M')}.{excel_ext}"

    if excel_ext == EXT_XLSM:
        st.success(
            f"Ready to export ✅  Reviewer: {st.session_state.reviewer_name} · "
            f"Reviewed at: {st.session_state.reviewed_at}  "
            f"(Macro-enabled: double-click 'Reviewed & OK' to toggle and stamp the date.)"
        )
    else:
        st.success(
            f"Ready to export ✅  Reviewer: {st.session_state.reviewer_name} · "
            f"Reviewed at: {st.session_state.reviewed_at}"
        )
        st.info("template.xlsm not found — exporting plain .xlsx without macros.")

    st.download_button(
        label="📥 Download FINAL Excel",
        data=excel_bytes,
        file_name=filename,
        mime=excel_mime,
        use_container_width=True,
    )
    audit("final_excel_download_ready", filename)

    st.markdown("</div>", unsafe_allow_html=True)


# =============================================================================
# Router
# =============================================================================
if st.session_state.view_step == "Setup":
    render_setup_step()
elif st.session_state.view_step == "Review":
    render_review_step()
elif st.session_state.view_step == "Export Excel":
    render_export_step()

st.markdown(f"<div class='footer-version'>Tool version: {TOOL_VERSION}</div>", unsafe_allow_html=True)
