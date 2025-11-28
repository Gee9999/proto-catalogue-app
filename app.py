import streamlit as st
import pandas as pd
import os
import re
from io import BytesIO
from PIL import Image
from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage

st.set_page_config(page_title="Proto Catalogue Builder", layout="wide")

st.title("📸 Proto Catalogue Builder")
st.write("Upload a price Excel file and a photo folder. I'll match photos to codes, descriptions and VAT-inclusive prices.")

# ---------------------------
# 1. HELPER: Identify columns
# ---------------------------

def find_column(columns, patterns):
    """
    Find the first column whose name matches any of the regex patterns (case-insensitive).
    """
    for col in columns:
        name = str(col).lower()
        for pat in patterns:
            if re.search(pat, name):
                return col
    return None

# -------------------------------------
# 2. Extract base numeric code from filename
# -------------------------------------

def extract_digits_from_filename(filename: str) -> str:
    """
    Extract the longest continuous run of digits from the filename.
    e.g. '8613900012-20PCS.jpg' -> '8613900012'
    """
    nums = re.findall(r"(\d+)", filename)
    if not nums:
        return None
    return max(nums, key=len)  # longest run of digits

# -----------------------------------------------------
# 3. String similarity: longest common *contiguous* substring
# -----------------------------------------------------

def longest_common_substring(a: str, b: str) -> int:
    """
    Return length of the longest common contiguous substring between a and b.
    e.g. a='8613900012', b='8613900032' -> 8 (the '86139000' part).
    """
    n, m = len(a), len(b)
    if n == 0 or m == 0:
        return 0
    best = 0
    # DP table
    dp = [[0] * (m + 1) for _ in range(n + 1)]
    for i in range(1, n + 1):
        ai = a[i - 1]
        for j in range(1, m + 1):
            if ai == b[j - 1]:
                dp[i][j] = dp[i - 1][j - 1] + 1
                if dp[i][j] > best:
                    best = dp[i][j]
            else:
                dp[i][j] = 0
    return best

# -----------------------------------------------------
# 4. Match photos to Excel codes by "closest number"
# -----------------------------------------------------

def match_photos_to_excel(df: pd.DataFrame, image_folder: str, min_score: float = 0.6) -> pd.DataFrame:
    """
    For each image in the folder, extract its numeric code, then find the closest
    matching CODE in the Excel file based on longest common substring ratio.
    min_score is the minimum similarity ratio to accept a match.
    """
    # Normalised numeric version of Excel codes (digits only)
    df = df.copy()
    df["CODE_STR"] = df["CODE"].astype(str)
    df["CODE_NUM"] = df["CODE_STR"].str.replace(r"\D", "", regex=True)

    images = [
        f for f in os.listdir(image_folder)
        if f.lower().endswith((".jpg", ".jpeg", ".png"))
    ]

    matched_rows = []

    for img_file in images:
        digits = extract_digits_from_filename(img_file)
        if not digits:
            continue

        # Step 1: try 6-digit prefix filter to reduce candidate set
        candidates = df[df["CODE_NUM"].str.startswith(digits[:6])]

        # Step 2: if none, fall back to 4-digit prefix
        if candidates.empty and len(digits) >= 4:
            candidates = df[df["CODE_NUM"].str.startswith(digits[:4])]

        # Step 3: if still none, use all rows as candidates
        if candidates.empty:
            candidates = df

        best_row = None
        best_score = 0.0

        for _, row in candidates.iterrows():
            code_num = str(row["CODE_NUM"])
            lcs_len = longest_common_substring(digits, code_num)
            denom = max(len(digits), len(code_num))
            score = lcs_len / denom if denom > 0 else 0.0

            if score > best_score:
                best_score = score
                best_row = row

        if best_row is not None and best_score >= min_score:
            matched_rows.append({
                "FILENAME": img_file,
                "PHOTO_CODE": digits,
                "CODE": best_row["CODE_STR"],
                "DESCRIPTION": best_row["DESCRIPTION"],
                "PRICE_INCL": best_row["PRICE_INCL"],
                "MATCH_SCORE": round(best_score, 3),
            })

    return pd.DataFrame(matched_rows)

# ---------------------------------------------------------
# 5. BUILD EXCEL WITH THUMBNAILS
# ---------------------------------------------------------

def build_excel_with_thumbnails(df: pd.DataFrame, image_folder: str) -> BytesIO:
    """
    Build an Excel file with:
      A: Photo thumbnail
      B: CODE (from Excel)
      C: DESCRIPTION
      D: PRICE_INCL
      E: FILENAME
      F: PHOTO_CODE (numeric extracted from filename)
      G: MATCH_SCORE
    """
    from openpyxl.utils import get_column_letter  # local import to keep top clean

    wb = Workbook()
    ws = wb.active
    ws.title = "Catalogue"

    headers = ["PHOTO", "CODE", "DESCRIPTION", "PRICE_INCL", "FILENAME", "PHOTO_CODE", "MATCH_SCORE"]
    ws.append(headers)

    row_index = 2

    for _, row in df.iterrows():
        img_path = os.path.join(image_folder, row["FILENAME"])

        # Insert image thumbnail, if present
        if os.path.exists(img_path):
            try:
                img = Image.open(img_path)
                img.thumbnail((150, 150))
                temp_io = BytesIO()
                img.save(temp_io, format="JPEG")
                temp_io.seek(0)
                xl_img = XLImage(temp_io)
                xl_img.anchor = f"A{row_index}"
                ws.add_image(xl_img)
            except Exception:
                pass

        ws[f"B{row_index}"] = row.get("CODE", "")
        ws[f"C{row_index}"] = row.get("DESCRIPTION", "")
        ws[f"D{row_index}"] = row.get("PRICE_INCL", "")
        ws[f"E{row_index}"] = row.get("FILENAME", "")
        ws[f"F{row_index}"] = row.get("PHOTO_CODE", "")
        ws[f"G{row_index}"] = row.get("MATCH_SCORE", "")

        ws.row_dimensions[row_index].height = 120
        row_index += 1

    # Column widths
    widths = {
        "A": 25,
        "B": 18,
        "C": 50,
        "D": 15,
        "E": 35,
        "F": 18,
        "G": 12,
    }
    for col, width in widths.items():
        ws.column_dimensions[col].width = width

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# -------------------------------------------------------------
# 6. STREAMLIT UI
# -------------------------------------------------------------

st.subheader("Step 1: Upload your price Excel file")
uploaded_excel = st.file_uploader(
    "Excel with CODE, DESCRIPTION, VAT-inclusive price",
    type=["xlsx", "xls"]
)

st.subheader("Step 2: Enter the folder path with your product photos")
uploaded_folder = st.text_input("Photo folder path (e.g. C:/Users/George/Pictures/beads)")

if uploaded_excel and uploaded_folder:
    if not os.path.isdir(uploaded_folder):
        st.error("The folder path you entered does not exist on this machine. Please check the path.")
    else:
        with st.spinner("Reading Excel and detecting columns…"):
            df = pd.read_excel(uploaded_excel)
            cols = list(df.columns)

            # Try to find columns generically
            col_code = find_column(cols, [r"^code$", r"code", r"item", r"product"])
            col_desc = find_column(cols, [r"desc", r"description", r"details"])
            col_price = find_column(cols, [r"incl", r"vat", r"price-a", r"price a", r"price_incl", r"selling"])

            missing = []
            if not col_code:
                missing.append("CODE")
            if not col_desc:
                missing.append("DESCRIPTION")
            if not col_price:
                missing.append("VAT-inclusive price")

            if missing:
                st.error(
                    "I couldn't find these columns in your Excel file: "
                    + ", ".join(missing)
                    + ".\n\nColumns I see are: "
                    + ", ".join(map(str, cols))
                )
            else:
                st.success(
                    f"Detected columns → CODE: '{col_code}', "
                    f"DESCRIPTION: '{col_desc}', PRICE_INCL: '{col_price}'"
                )

                df_norm = df.rename(columns={
                    col_code: "CODE",
                    col_desc: "DESCRIPTION",
                    col_price: "PRICE_INCL",
                })

                with st.spinner("Matching photos to closest codes and prices…"):
                    matched_df = match_photos_to_excel(df_norm, uploaded_folder, min_score=0.6)

                if matched_df.empty:
                    st.warning("No matches were found. Try checking filenames/Excel codes or lowering the match threshold in the code.")
                else:
                    st.write("### ✅ Matched items (best numeric match per photo)")
                    st.dataframe(matched_df)

                    excel_file = build_excel_with_thumbnails(matched_df, uploaded_folder)

                    st.download_button(
                        "📥 Download matched catalogue Excel",
                        data=excel_file,
                        file_name="catalogue_matched.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )
