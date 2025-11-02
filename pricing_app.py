import streamlit as st
import pandas as pd
import pdfplumber
from datetime import date
from io import BytesIO

# -----------------------------------------------------------
# CONFIG
# -----------------------------------------------------------
st.set_page_config(page_title="Jomar Contract Pricing Applier", layout="wide")

# Path to your standardized workbook *inside the repo*
PRODUCTS_PATH = "JomarList_10272025.xlsx"   # <-- adjust if needed
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
# HELPERS
# -----------------------------------------------------------
@st.cache_data
def load_product_workbook(path: str):
    """
    Load the standardized Excel from the repo.
    Must contain:
      - Flat List
      - Model Group
    """
    xls = pd.ExcelFile(path)
    flat = pd.read_excel(xls, sheet_name=FLAT_SHEET_NAME)
    model = pd.read_excel(xls, sheet_name=GROUP_SHEET_NAME)
    return flat, model


def extract_contract_from_pdf(pdf_file) -> pd.DataFrame:
    """
    Read a PDF like the sample you shared and return ONLY the rows under
    the header:
      Product / Group / Line | Code | Start Date | End Date | Price / Multi
    We ignore all the branch / header text above it.
    """
    rows = []
    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables()
            for table in tables:
                keep = False
                for r in table:
                    # r is a list of cells
                    cells = [(c.strip() if isinstance(c, str) else c) for c in r]
                    if not any(cells):
                        continue

                    # detect header row
                    if cells[0] and HEADER_MARKER in cells[0]:
                        keep = True
                        continue

                    # after header we start capturing rows
                    if keep:
                        rows.append(cells)

    if not rows:
        # return empty in the normalized shape
        return pd.DataFrame(
            columns=["key_value", "key_type", "start_date", "end_date", "multiplier"]
        )

    # your PDF had 6 columns (the 6th is that trailing X)
    # but if a page is a bit off, we'll pad
    max_len = max(len(r) for r in rows)
    norm_rows = [r + [None] * (max_len - len(r)) for r in rows]

    # try to build with the expected columns
    # we'll name the first 6 like this; extra cols get dropped
    col_names = [
        "Product / Group / Line",
        "Code",
        "Start Date",
        "End Date",
        "Price / Multi",
        "Extra",
    ][:max_len]

    df_raw = pd.DataFrame(norm_rows, columns=col_names)

    # normalize column names to what we use later
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
    "This app uses a **standard Excel** from the repo. "
    "Upload a customer's **contract PDF** and we'll apply multipliers to the standard list."
)

# load standard workbook
try:
    flat_list, model_group = load_product_workbook(PRODUCTS_PATH)
except FileNotFoundError:
    st.error(f"⚠️ Could not find standardized Excel at `{PRODUCTS_PATH}`. Make sure it's in the repo.")
    st.stop()

# merge model info onto flat list so each part knows its Sub-Group / Line / Sub-Line
flat_merged = flat_list.merge(
    model_group[["Part #", "Sub-Group", "Line", "Sub-Line"]],
    on="Part #",
    how="left"
)

st.subheader("📦 Standard product master (preview)")
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
        excel_bytes = to_excel_bytes({"Jomar List Pricing": priced_df})
        st.download_button(
            label="⬇️ Download priced Excel",
            data=excel_bytes,
            file_name="priced_flat_list.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
else:
    st.info("Upload a contract PDF to apply multipliers.")


