# sites/core_pdf.py
from typing import Callable, Dict, Optional
import pandas as pd
import fitz

def extract_generic(
    file_bytes: bytes,
    file_name: str,
    site_config: Dict,
) -> pd.DataFrame:
    """
    Returns a dataframe in the app schema:
    process_step, chapter, step, parameter, theoretical_value, process_instruction, page
    """
    # 1) write bytes to temp file or open with fitz from stream
    doc = fitz.open(stream=file_bytes, filetype="pdf")

    # 2) run your existing logic:
    #    detect stages, extract rows, compute Data tag etc.
    # 3) map outputs into the app schema columns

    # Example placeholder:
    rows = []
    # ... fill rows ...
    df = pd.DataFrame(rows)

    return df
