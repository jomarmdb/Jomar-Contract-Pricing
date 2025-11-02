import streamlit as st
import pandas as pd
import pdfplumber
from datetime import date
from io import BytesIO
from pathlib import Path

# -----------------------------------------------------------
# CONFIG
# -----------------------------------------------------------
st.set_page_config(page_title="Jomar Contract Pricing Applier", layout="wide")

BASE_DIR = Path(__file__).parent

# put your standardized workbook in the same folder as this script
PRODUCTS_PATH = BASE_DIR / "JomarList_10272025.xlsx"   # <-- make sure this filename matches exactly

FLAT_SHEET_NAME = "Jomar List Pricing"
GROUP_SHEET_NAME = "Model Group"

# your Excel headers start on row 9 (1-based) -> header=8 (0-based)
HEADER_ROW_INDEX = 8

# PDF header text we look for (but we will detect it by words now)
HEADER_MARKER = "Product"

# PDF code → our internal type
CODE_MAP = {
    "P": "PART",
    "U": "SUBLINE",
    "S": "SUBGROUP",
    "L": "LINE",
    "G": None,  # ignore group-level contracts
}

# -----------------------------------------------------------
# HELPER FUNCTIONS (Excel)
# -----------------------------------------------------------

def normalize_cols(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = df.columns.str.strip()
    return df

def normalize_flat(df: pd.DataFrame) -> pd.DataFrame:
    """
    Make sure 'Jomar List Pricing' has at least:
      - Part #
      - List Price
    even if the actual headers are a little different.
    """
    df = normalize_cols(df)
    rename_map = {}

    # part # variants
    if "Part #" not in df.columns:
        if "Part#" in df.columns:
            rename_map["Part#"] = "Part #"
        elif "Part Number" in df.columns:
            rename_map["Part Number"] = "Part #"
        elif "Part No" in df.columns:
            rename_map["Part No"] = "Part #"

    # list price variants
    if "List Price" not in df.columns:
        if "List" in df.columns:
            rename_map["List"] = "List Price"
        elif "Price" in df.columns:
            rename_map["Price"] = "List Price"

    return df.rename(columns=rename_map)

def normalize_model(df: pd.DataFrame) -> pd.DataFrame:
    """
    Make sure 'Model Group' has:
      - Part #
      - Sub-Group
      - Line
      - Sub-Line
    We'll ignore 'Model #' and 'Group'.
    """
    df = normalize_cols(df)
    rename_map = {}

    # part #
    if "Part #" not in df.columns and "Part#" in df.columns:
        rename_map["Part#"] = "Part #"

    # sub-group
    if "Sub-Group" not in df.columns:
        if "Sub Group" in df.columns:
            rename_map["Sub Group"] = "Sub-Group"
        elif "Subgroup" in df.columns:
            rename_map["Subgroup"] = "Sub-Group"

    # sub-line
    if "Sub-Line" not in df.columns:
        if "Sub Line" in df.columns:
            rename_map["Sub Line"] = "Sub-Line"
        elif "Subline" in df.columns:
            rename_map["Subline"] = "Sub-Line"

    # line (in case there's a trailing space)
    if "Line" not in df.columns and "Line " in df.columns:
        rename_map["Line "] = "Line"

    return df.rename(columns=rename_map)

@st.cache_data
def load_product_workbook(path: Path):
    """
    Load your standardized workbook.
    We read both sheets using header=8 because your titles start on row 9.
    """
    xls = pd.ExcelFile(path)
    flat = pd.read_excel(xls, sheet_name=FLAT_SHEET_NAME, header=HEADER_ROW_INDEX)
    model = pd.read_excel(xls, sheet_name=GROUP_SHEET_NAME, header=HEADER_ROW_INDEX)
    return flat, model

# -----------------------------------------------------------
# PDF PARSER (word-based for your actual PDF)
# -----------------------------------------------------------

def extract_contr_
