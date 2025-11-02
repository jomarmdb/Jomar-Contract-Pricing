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

def norm_key(val):
    """
    Normalize strings so PDF keys and Excel keys can match.
    - turn to string
    - strip spaces
    - uppercase
    - normalize weird dashes and quotes
    """
    if pd.isna(val):
        return None
    s = str(val).strip()
    if not s:
        return None
    s = (s
         .replace("–", "-")
         .replace("—", "-")
         .replace("-", "-")   # non-breaking hyphen
         .replace("’", "'"))
    return s.upper()

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

    # list price variants (this is just a safety net)
    if "List Price" not in df.columns:
        for col in df.columns:
            if "List Price" in str(col):
                rename_map[col] = "List Price"
                break
        else:
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
    Pricing sheet: header row is 9 → header=8
    Model sheet: looks like your columns start right away → header=0
    """
    xls = pd.ExcelFile(path)
    flat = pd.read_excel(xls, sheet_name=FLAT_SHEET_NAME, header=HEADER_ROW_INDEX)
    model = pd.read_excel(xls, sheet_name=GROUP_SHEET_NAME, header=0)
    return flat, model
    

# -----------------------------------------------------------
# PDF PARSER (word-based)
# -----------------------------------------------------------

def extract_contract_from_pdf(pdf_file) -> pd.DataFrame:
    rows = []

    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            words = page.extract_words()
            line_dict = {}
            for w in words:
                top = round(w["top"])
                line_dict.setdefault(top, []).append(w)

            sorted_tops = sorted(line_dict.keys())
            header_seen = False

            for top in sorted_tops:
                ws = sorted(line_dict[top], key=lambda x: x["x0"])
                texts = [w["text"] for w in ws]

                # detect the header-ish line
                if not header_seen:
                    if ("Product" in texts and "Group" in texts and "Line" in texts):
                        header_seen = True
                    continue

                # after header: slice by x0 ranges (approx from your sample)
                product_parts = [w["text"] for w in ws if w["x0"] < 200]
                code_parts    = [w["text"] for w in ws if 200 <= w["x0"] < 250]
                start_parts   = [w["text"] for w in ws if 250 <= w["x0"] < 305]
                end_parts     = [w["text"] for w in ws if 300 <= w["x0"] < 370]
                multi_parts   = [w["text"] for w in ws if 370 <= w["x0"] < 430]

                if not product_parts:
                    continue

                product = " ".join(product_parts).strip()
                code    = code_parts[0] if code_parts else ""
                start   = start_parts[0] if start_parts else ""
                end     = end_parts[0] if end_parts else ""
                multi   = multi_parts[0] if multi_parts else ""

                if code not in ("P", "U", "S", "L", "G"):
                    continue

                rows.append((product, code, start, end, multi))

    if not rows:
        return pd.DataFrame(
            columns=["key_value", "key_type", "start_date", "end_date", "multiplier"]
        )

    df = pd.DataFrame(rows, columns=["key_value", "code", "start_date", "end_date", "multiplier"])
    df["key_type"] = df["code"].map(CODE_MAP)
    df = df[df["key_type"].notna()]

    df["start_date"] = pd.to_datetime(df["start_date"], errors="coerce")
    df["end_date"]   = pd.to_datetime(df["end_date"], errors="coerce")
    df["multiplier"] = pd.to_numeric(df["multiplier"], errors="coerce")

    df["key_norm"] = df["key_value"].apply(norm_key)
    return df[["key_value", "key_type", "start_date", "end_date", "multiplier", "key_norm"]]


# -----------------------------------------------------------
# PRICING LOGIC
# -----------------------------------------------------------

def filter_active(contract_df: pd.DataFrame, as_of: date | None = None) -> pd.DataFrame:
    """
    Keep only rows that are 'active' as of today.

    Special rule for your PDFs:
    - If End Date is something ancient (year < 2000, e.g. 12/31/1949),
      we will treat it as 'no end date' / still active.
    """
    if as_of is None:
        as_of = date.today()

    def _active(r):
        # start check
        start_ok = pd.isna(r["start_date"]) or (r["start_date"].date() <= as_of)

        # end check
        if pd.isna(r["end_date"]):
            end_ok = True
        else:
            end_year = r["end_date"].year
            if end_year < 2000:
                # <-- THIS is the important part for your 12/31/1949 rows
                end_ok = True
            else:
                end_ok = r["end_date"].date() >= as_of

        return start_ok and end_ok

    return contract_df[contract_df.apply(_active, axis=1)]

def apply_contract(
    flat_df: pd.DataFrame,
    contract_df: pd.DataFrame,
    default_mult: float = 0.50,
    list_price_col: str = "List Price",
) -> pd.DataFrame:
    active = filter_active(contract_df).copy()

    # build lookup dicts by type (normalized)
    part_map = {}
    subline_map = {}
    subgroup_map = {}
    line_map = {}

    for _, r in active.iterrows():
        key = r.get("key_norm")
        if not key:
            continue
        mult = r["multiplier"]
        if r["key_type"] == "PART":
            part_map[key] = mult
        elif r["key_type"] == "SUBLINE":
            subline_map[key] = mult
        elif r["key_type"] == "SUBGROUP":
            subgroup_map[key] = mult
        elif r["key_type"] == "LINE":
            line_map[key] = mult

    # make sure target columns exist
    if "Multiplier" not in flat_df.columns:
        flat_df["Multiplier"] = None
    if "Net Price" not in flat_df.columns:
        flat_df["Net Price"] = None

    # make sure list price is numeric
    flat_df[list_price_col] = pd.to_numeric(flat_df[list_price_col], errors="coerce")

    multipliers = []
    sources = []

    for _, row in flat_df.iterrows():
        part     = norm_key(row.get("Part #"))
        subline  = norm_key(row.get("Sub-Line"))
        subgroup = norm_key(row.get("Sub-Group"))
        line     = norm_key(row.get("Line"))

        # 1) exact part
        if part and part in part_map:
            m = float(part_map[part])
            multipliers.append(m)
            sources.append(f"PART:{part}")
            continue

        # 2) sub-line
        if subline and subline in subline_map:
            m = float(subline_map[subline])
            multipliers.append(m)
            sources.append(f"SUBLINE:{subline}")
            continue

        # 3) sub-group
        if subgroup and subgroup in subgroup_map:
            m = float(subgroup_map[subgroup])
            multipliers.append(m)
            sources.append(f"SUBGROUP:{subgroup}")
            continue

        # 4) line
        if line and line in line_map:
            m = float(line_map[line])
            multipliers.append(m)
            sources.append(f"LINE:{line}")
            continue

        # 5) default
        multipliers.append(default_mult)
        sources.append("DEFAULT:0.50")

    # write into your existing columns
    flat_df["Multiplier"] = multipliers
    flat_df["Net Price"] = flat_df[list_price_col] * flat_df["Multiplier"]

    # keep audit
    flat_df["Match Source"] = sources

    return flat_df

def to_excel_bytes(df_dict: dict[str, pd.DataFrame]) -> bytes:
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
    "This app loads your **standardized Excel** (headers on row 9), "
    "parses the contract PDF you upload, and applies multipliers in the order: "
    "**Part → Sub-Line → Sub-Group → Line → 0.50**."
)

# load workbook
try:
    flat_list, model_group = load_product_workbook(PRODUCTS_PATH)
except FileNotFoundError:
    st.error(f"⚠️ Could not find standardized Excel at `{PRODUCTS_PATH}`.")
    st.stop()

# normalize columns
flat_list = normalize_flat(flat_list)
model_group = normalize_model(model_group)

# show for debugging
st.write("📄 Jomar List Pricing columns:", list(flat_list.columns))
st.write("📄 Model Group columns:", list(model_group.columns))

# check required columns
needed_model_cols = ["Part #", "Sub-Group", "Line", "Sub-Line"]
missing = [c for c in needed_model_cols if c not in model_group.columns]
if missing:
    st.error(f"'Model Group' sheet is missing these columns: {missing}")
    st.stop()

if "Part #" not in flat_list.columns:
    st.error("'Jomar List Pricing' sheet is missing 'Part #' column.")
    st.stop()

# merge model columns onto pricing
flat_merged = flat_list.merge(
    model_group[["Part #", "Sub-Group", "Line", "Sub-Line"]],
    on="Part #",
    how="left"
)

# 🔎 detect the actual list-price column name AFTER merge
list_price_col = None
for col in flat_merged.columns:
    if "List Price" in str(col):
        list_price_col = col
        break

st.write("🔎 Detected list-price column (post-merge):", list_price_col)

if list_price_col is None:
    st.error("Could not find a column that contains 'List Price' in the merged pricing sheet.")
    st.stop()

st.subheader("📦 Standard product master (merged preview)")
st.dataframe(flat_merged.head(25))

# upload contract PDF
pdf_file = st.file_uploader("📄 Upload contract PDF", type=["pdf"])

if pdf_file is not None:
    contract_df = extract_contract_from_pdf(pdf_file)

    st.subheader("🧾 Parsed contract rows")
    st.dataframe(contract_df)

    if contract_df.empty:
        st.warning("No contract rows were found under the header. Check the PDF format.")
    else:
        priced_df = apply_contract(
            flat_merged.copy(),
            contract_df,
            default_mult=0.50,
            list_price_col=list_price_col,   # 👈 use detected name
        )

        st.subheader("💰 Priced output (first 100 rows)")
        st.dataframe(priced_df.head(100))

        excel_bytes = to_excel_bytes({"Jomar List Pricing (Priced)": priced_df})

        st.download_button(
            label="⬇️ Download priced Excel",
            data=excel_bytes,
            file_name="priced_jomar_list.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
else:
    st.info("Upload a contract PDF to apply multipliers.")




