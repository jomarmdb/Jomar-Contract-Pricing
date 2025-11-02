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

# base directory of this script
BASE_DIR = Path(__file__).parent

# Path to your standardized workbook *inside the repo*
# put the file next to this script, or adjust to BASE_DIR / "data" / ...
PRODUCTS_PATH = BASE_DIR / "JomarList_10272025.xlsx"   # <-- change to your exact filename

# Sheet names inside that workbook
FLAT_SHEET_NAME = "Jomar List Pricing"
GROUP_SHEET_NAME = "Model Group"

# PDF header text we look for
HEADER_MARKER = "Product / Group / Line"

# PDF code → our type
CODE_MAP = {
    "P": "PART",
    "U": "SUBLINE",
    "S": "SUBGROUP",
    "L": "LINE",
    "G": None,  # we ignore group-level contracts
}

# -----------------------------------------------------------
# HELPER FUNCTIONS
# -----------------------------------------------------------

def normalize_cols(df: pd.DataFrame) -> pd.DataFrame:
    """Strip whitespace from column names."""
    df = df.copy()
    df.columns = df.columns.str.strip()
    return df


def normalize_flat(df: pd.DataFrame) -> pd.DataFrame:
    """
    Make sure the pricing sheet (Jomar List Pricing) has:
      - Part #
      - List Price
    even if the user called them Part#, Part Number, List, etc.
    """
    df = normalize_cols(df)
    rename_map = {}

    # Part #
    if "Part #" not in df.columns:
        if "Part#" in df.columns:
            rename_map["Part#"] = "Part #"
        elif "Part Number" in df.columns:
            rename_map["Part Number"] = "Part #"
        elif "Part No" in df.columns:
            rename_map["Part No"] = "Part #"

    # List Price
    if "List Price" not in df.columns:
        if "List" in df.columns:
            rename_map["List"] = "List Price"
        elif "Price" in df.columns:
            rename_map["Price"] = "List Price"

    return df.rename(columns=rename_map)


def normalize_model(df: pd.DataFrame) -> pd.DataFrame:
    """
    Make sure the Model Group sheet has:
      - Part #
      - Sub-Group
      - Line
      - Sub-Line

    and ignore "Model #" and "Group", since you said those
    are just leftover / not used for contracting.
    """
    df = normalize_cols(df)
    rename_map = {}

    # Part #
    if "Part #" not in df.columns and "Part#" in df.columns:
        rename_map["Part#"] = "Part #"

    # Sub-Group variants
    if "Sub-Group" not in df.columns:
        if "Sub Group" in df.columns:
            rename_map["Sub Group"] = "Sub-Group"
        elif "Subgroup" in df.columns:
            rename_map["Subgroup"] = "Sub-Group"

    # Sub-Line variants
    if "Sub-Line" not in df.columns:
        if "Sub Line" in df.columns:
            rename_map["Sub Line"] = "Sub-Line"
        elif "Subline" in df.columns:
            rename_map["Subline"] = "Sub-Line"

    # Line variants (common: stray space)
    if "Line" not in df.columns and "Line " in df.columns:
        rename_map["Line "] = "Line"

    return df.rename(columns=rename_map)


@st.cache_data
def load_product_workbook(path: Path):
    """
    Load the standardized Excel from the repo.
    Must contain:
      - Jomar List Pricing
      - Model Group
    """
    xls = pd.ExcelFile(path)
    xls = pd.ExcelFile(path)
    # Skip the first 8 rows so pandas treats row 9 as the header row
    flat = pd.read_excel(xls, sheet_name=FLAT_SHEET_NAME, header=8)
    model = pd.read_excel(xls, sheet_name=GROUP_SHEET_NAME, header=0)
    return flat, model


def extract_contract_from_pdf(pdf_file) -> pd.DataFrame:
    """
    Read a PDF like the sample and return ONLY the rows under
    the header:
      Product / Group / Line | Code | Start Date | End Date | Price / Multi
    We ignore all the customer/branch stuff at the top.
    """
    rows = []
    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables()
            for table in tables:
                keep = False
                for r in table:
                    # normalize each cell
                    cells = [(c.strip() if isinstance(c, str) else c) for c in r]
                    if not any(cells):
                        continue

                    # detect header
                    if cells[0] and HEADER_MARKER in cells[0]:
                        keep = True
                        continue

                    # only capture rows after header
                    if keep:
                        rows.append(cells)

    if not rows:
        # return empty in normalized shape
        return pd.DataFrame(
            columns=["key_value", "key_type", "start_date", "end_date", "multiplier"]
        )

    # some pages might have 5 cols, some 6 -> normalize
    max_len = max(len(r) for r in rows)
    norm_rows = [r + [None] * (max_len - len(r)) for r in rows]

    # expected columns (your PDF had 6, last was an X)
    col_names = [
        "Product / Group / Line",
        "Code",
        "Start Date",
        "End Date",
        "Price / Multi",
        "Extra",
    ][:max_len]

    df_raw = pd.DataFrame(norm_rows, columns=col_names)

    # normalize column names to our internal schema
    df = df_raw.rename(
        columns={
            "Product / Group / Line": "key_value",
            "Code": "code",
            "Start Date": "start_date",
            "End Date": "end_date",
            "Price / Multi": "multiplier",
        }
    ).copy()

    # map code -> key_type
    df["key_type"] = df["code"].map(CODE_MAP)

    # drop unsupported rows (G or unknown)
    df = df[df["key_type"].notna()]

    # parse dates
    df["start_date"] = pd.to_datetime(df["start_date"], errors="coerce")
    df["end_date"] = pd.to_datetime(df["end_date"], errors="coerce")

    # numeric multiplier
    df["multiplier"] = pd.to_numeric(df["multiplier"], errors="coerce")

    return df[["key_value", "key_type", "start_date", "end_date", "multiplier"]]


def filter_active(contract_df: pd.DataFrame, as_of: date | None = None) -> pd.DataFrame:
    """Return only rows where start_date <= today <= end_date (if present)."""
    if as_of is None:
        as_of = date.today()

    def _active(r):
        start_ok = pd.isna(r["start_date"]) or (r["start_date"].date() <= as_of)
        end_ok = pd.isna(r["end_date"]) or (r["end_date"].date() >= as_of)
        return start_ok and end_ok

    return contract_df[contract_df.apply(_active, axis=1)]


def apply_contract(flat_df: pd.DataFrame, contract_df: pd.DataFrame, default_mult: float = 0.50) -> pd.DataFrame:
    """
    Apply the pricing priority:
      PART (P) → SUBLINE (U) → SUBGROUP (S) → LINE (L) → default.
    Assumes flat_df already has: Part #, Sub-Line, Sub-Group, Line, List Price
    """
    active = filter_active(contract_df)

    # make sure list price is numeric
    flat_df["List Price"] = pd.to_numeric(flat_df["List Price"], errors="coerce")

    multipliers = []
    sources = []

    for _, row in flat_df.iterrows():
        part = row.get("Part #")
        subline = row.get("Sub-Line")
        subgroup = row.get("Sub-Group")
        line = row.get("Line")

        # 1. PART
        hit = active[
            (active["key_type"] == "PART") & (active["key_value"] == part)
        ]
        if not hit.empty:
            m = float(hit.iloc[0]["multiplier"])
            multipliers.append(m)
            sources.append(f"PART:{part}")
            continue

        # 2. SUBLINE
        if pd.notna(subline):
            hit = active[
                (active["key_type"] == "SUBLINE") & (active["key_value"] == subline)
            ]
            if not hit.empty:
                m = float(hit.iloc[0]["multiplier"])
                multipliers.append(m)
                sources.append(f"SUBLINE:{subline}")
                continue

        # 3. SUBGROUP
        if pd.notna(subgroup):
            hit = active[
                (active["key_type"] == "SUBGROUP") & (active["key_value"] == subgroup)
            ]
            if not hit.empty:
                m = float(hit.iloc[0]["multiplier"])
                multipliers.append(m)
                sources.append(f"SUBGROUP:{subgroup}")
                continue

        # 4. LINE
        if pd.notna(line):
            hit = active[
                (active["key_type"] == "LINE") & (active["key_value"] == line)
            ]
            if not hit.empty:
                m = float(hit.iloc[0]["multiplier"])
                multipliers.append(m)
                sources.append(f"LINE:{line}")
                continue

        # 5. default
        multipliers.append(default_mult)
        sources.append("DEFAULT:0.50")

    flat_df["Contract Multiplier"] = multipliers
    flat_df["Match Source"] = sources
    flat_df["Contract Net Price"] = flat_df["List Price"] * flat_df["Contract Multiplier"]

    return flat_df


def to_excel_bytes(df_dict: dict[str, pd.DataFrame]) -> bytes:
    """Write one or more DataFrames to an in-memory Excel file."""
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for sheet_name, df in df_dict.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)
    output.seek(0)
    return output.getvalue()


# -----------------------------------------------------------
# UI
# -----------------------------------------------------------
st.title("Jomar Contract Pricing Applier")

st.write(
    "This app uses your **standardized Excel** in the repo "
    f"(`{PRODUCTS_PATH.name}`) and applies multipliers from a customer PDF. "
    "Priority: **Part → Sub-Line → Sub-Group → Line → 0.50**."
)

# load standard workbook
try:
    flat_list, model_group = load_product_workbook(PRODUCTS_PATH)
except FileNotFoundError:
    st.error(f"⚠️ Could not find standardized Excel at `{PRODUCTS_PATH}`.")
    st.stop()

# normalize column names on BOTH sheets
flat_list = normalize_flat(flat_list)
model_group = normalize_model(model_group)

# show what we actually have (super helpful while setting up)
st.write("📄 Jomar List Pricing columns:", list(flat_list.columns))
st.write("📄 Model Group columns:", list(model_group.columns))

# make sure required columns are there
needed_model_cols = ["Part #", "Sub-Group", "Line", "Sub-Line"]
missing = [c for c in needed_model_cols if c not in model_group.columns]
if missing:
    st.error(f"Model Group sheet is missing these columns: {missing}")
    st.stop()

if "Part #" not in flat_list.columns:
    st.error("Pricing sheet is missing 'Part #' column.")
    st.stop()

# now it's safe to merge
flat_merged = flat_list.merge(
    model_group[["Part #", "Sub-Group", "Line", "Sub-Line"]],
    on="Part #",
    how="left"
)

st.subheader("📦 Standard product master (merged preview)")
st.dataframe(flat_merged.head(25))

# file uploader for contract PDF
pdf_file = st.file_uploader("📄 Upload contract PDF", type=["pdf"])

if pdf_file is not None:
    # 1) parse the PDF
    contract_df = extract_contract_from_pdf(pdf_file)

    st.subheader("🧾 Parsed contract rows (from PDF)")
    st.dataframe(contract_df)

    if contract_df.empty:
        st.warning("No contract rows were found under the header. Check the PDF format.")
    else:
        # 2) apply contract to the standard product list
        priced_df = apply_contract(flat_merged.copy(), contract_df, default_mult=0.50)

        st.subheader("💰 Priced output (first 100 rows)")
        st.dataframe(priced_df.head(100))

        # 3) let user download the priced Excel
        excel_bytes = to_excel_bytes({"Jomar List Pricing (Priced)": priced_df})
        st.download_button(
            label="⬇️ Download priced Excel",
            data=excel_bytes,
            file_name="priced_jomar_list.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
else:
    st.info("Upload a contract PDF to apply multipliers.")


