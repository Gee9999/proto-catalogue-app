import streamlit as st
import pandas as pd
import os
import re
import tempfile
import io

from PIL import Image as PILImage
from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage


# ---------- Helpers for Excel price file ----------

def normalize_col_name(name: str) -> str:
    """
    Normalize column names to match things like:
    'PRICE-A INCL.', 'PRICE-A_x000a_INCL.', 'Description     ', etc.
    Keep only letters and numbers, lowercased.
    """
    if not isinstance(name, str):
        name = str(name)
    name = name.strip().lower()
    # keep only letters and digits
    return "".join(ch for ch in name if ch.isalnum())


def clean_code(value) -> str | None:
    """
    Turn any code cell into a plain digit string.
    - Handles floats like '8610400003.0'
    - Removes spaces and non-digits
    """
    if pd.isna(value):
        return None
    s = str(value).strip()
    # If it's like '8610400003.0', strip the .0
    if re.fullmatch(r"\d+\.0", s):
        s = s.split(".")[0]
    # keep only digits
    digits = re.sub(r"\D", "", s)
    return digits or None


def load_price_dataframe(uploaded_file) -> pd.DataFrame:
    """
    Load the uploaded Excel and extract a clean DataFrame with columns:
    CODE, DESCRIPTION, PRICE_A_INCL, CODE_STR
    """
    # Read everything as string so we don't lose codes
    df = pd.read_excel(uploaded_file, dtype=str)

    if df.empty:
        raise ValueError("The price Excel file appears to be empty.")

    norm_map = {col: normalize_col_name(col) for col in df.columns}

    # --- find CODE column ---
    code_col = None
    for col, norm in norm_map.items():
        if norm.startswith("code") or norm in {"itemcode", "productcode", "stockcode"}:
            code_col = col
            break
    if code_col is None:
        raise ValueError(
            "Could not find a CODE column. "
            f"Columns I see: {list(df.columns)}"
        )

    # --- find DESCRIPTION column ---
    desc_col = None
    for col, norm in norm_map.items():
        if norm.startswith("description") or norm in {"desc", "itemdescription", "productdescription"}:
            desc_col = col
            break
    if desc_col is None:
        raise ValueError(
            "Could not find a DESCRIPTION column. "
            f"Columns I see: {list(df.columns)}"
        )

    # --- find VAT-inclusive PRICE column ---
    price_col = None
    for col, norm in norm_map.items():
        # Examples we want to catch: 'priceaincl', 'priceavincl', 'pricevatincl' etc.
        if "price" in norm and "incl" in norm:
            price_col = col
            break
    if price_col is None:
        # last resort: any 'price' column
        for col, norm in norm_map.items():
            if "price" in norm:
                price_col = col
                break

    if price_col is None:
        raise ValueError(
            "Could not find a VAT-inclusive price column. "
            f"Columns I see: {list(df.columns)}"
        )

    # Build a smaller, clean DataFrame
    price_df = df[[code_col, desc_col, price_col]].copy()
    price_df.columns = ["CODE", "DESCRIPTION", "PRICE_A_INCL"]

    # Clean up codes
    price_df["CODE_STR"] = price_df["CODE"].apply(clean_code)

    # Drop rows without usable code
    price_df = price_df.dropna(subset=["CODE_STR"])

    # Remove duplicate codes (keep first)
    price_df = price_df.drop_duplicates(subset=["CODE_STR"], keep="first")

    return price_df


def build_price_map(price_df: pd.DataFrame) -> dict:
    """
    Build a dict mapping CODE_STR -> {CODE, DESCRIPTION, PRICE_A_INCL}
    """
    price_map = {}
    for _, row in price_df.iterrows():
        key = row["CODE_STR"]
        price_map[key] = {
            "CODE": row["CODE"],
            "DESCRIPTION": row["DESCRIPTION"],
            "PRICE_A_INCL": row["PRICE_A_INCL"],
        }
    return price_map


# ---------- Match photos to price list ----------

def extract_code_from_filename(filename: str) -> str | None:
    """
    Extract the first continuous block of digits from the filename.
    e.g. '8610400003-50pcs.JPG' -> '8610400003'
    """
    m = re.search(r"\d+", filename)
    if not m:
        return None
    return m.group(0)


def match_photos_to_prices(photo_files, price_df: pd.DataFrame):
    """
    Save photos to a temp directory, match them to price list by code,
    and return a DataFrame plus the temp directory path.
    """
    price_map = build_price_map(price_df)
    temp_dir = tempfile.mkdtemp()

    records = []

    for uploaded in photo_files:
        filename = uploaded.name
        code_from_photo = extract_code_from_filename(filename)

        # Save image file to temp_dir (so openpyxl can embed it)
        img_path = os.path.join(temp_dir, filename)
        with open(img_path, "wb") as f:
            f.write(uploaded.read())

        matched = price_map.get(code_from_photo) if code_from_photo else None

        record = {
            "PHOTO_FILE": filename,
            "CODE_FROM_PHOTO": code_from_photo or "",
            "CODE": matched["CODE"] if matched else "",
            "DESCRIPTION": matched["DESCRIPTION"] if matched else "",
            "PRICE_A_INCL": matched["PRICE_A_INCL"] if matched else "",
        }
        records.append(record)

    matched_df = pd.DataFrame(records)
    return matched_df, temp_dir


# ---------- Build Excel with thumbnails ----------

def build_excel_with_thumbnails(matched_df: pd.DataFrame, temp_dir: str) -> io.BytesIO:
    """
    Create an Excel file with:
    - Column A: Photo thumbnail
    - Column B: CODE
    - Column C: DESCRIPTION
    - Column D: PRICE-A INCL.

    Thumbnails are about 60x60 pixels and sit UNDER the row (no overlapping text).
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "Catalogue"

    # Column widths
    ws.column_dimensions["A"].width = 14  # image column
    ws.column_dimensions["B"].width = 14  # code
    ws.column_dimensions["C"].width = 45  # description
    ws.column_dimensions["D"].width = 14  # price

    # Header
    ws.append(["PHOTO", "CODE", "DESCRIPTION", "PRICE-A INCL."])

    for idx, row in matched_df.iterrows():
        excel_row = idx + 2  # header is row 1
        ws.row_dimensions[excel_row].height = 55  # a bit taller for the thumbnail

        code = row.get("CODE") or row.get("CODE_FROM_PHOTO") or ""
        desc = row.get("DESCRIPTION", "")
        price = row.get("PRICE_A_INCL", "")

        # Write text cells (all UNDER the image in the same row)
        ws.cell(row=excel_row, column=2, value=code)
        ws.cell(row=excel_row, column=3, value=desc)
        ws.cell(row=excel_row, column=4, value=price)

        photo_name = row["PHOTO_FILE"]
        img_path = os.path.join(temp_dir, photo_name)

        if os.path.exists(img_path):
            try:
                pil_img = PILImage.open(img_path)
                # Make a small thumbnail
                pil_img.thumbnail((60, 60))
                bio = io.BytesIO()
                pil_img.save(bio, format="PNG")
                bio.seek(0)
                bio.name = "thumb.png"

                xl_img = XLImage(bio)
                # Anchor image at the PHOTO cell in this row
                xl_img.anchor = f"A{excel_row}"
                ws.add_image(xl_img)
            except Exception:
                # If image embedding fails for some reason, we just skip the picture
                pass

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output


# ---------- Streamlit app ----------

def main():
    st.set_page_config(page_title="Photo Catalogue Builder", layout="wide")

    st.title("Photo Catalogue Builder")

    st.write(
        """
        1. Upload your **price Excel file** (any layout, as long as it has a code, description, and VAT-inclusive price column).  
        2. Upload the **product photos** (filenames must contain the product code, e.g. `8610400003-50pcs.JPG`).  
        3. Click **Generate Excel catalogue** to get an Excel file with **thumbnails + CODE + DESCRIPTION + PRICE-A INCL.**
        """
    )

    price_file = st.file_uploader(
        "Upload price Excel (e.g. PRODUCT DETAILS - BY CODE.xlsx)", 
        type=["xls", "xlsx"]
    )

    photo_files = st.file_uploader(
        "Upload product photos", 
        type=["jpg", "jpeg", "png"], 
        accept_multiple_files=True
    )

    if st.button("Generate Excel catalogue"):
        if not price_file:
            st.error("Please upload a price Excel file first.")
            return
        if not photo_files:
            st.error("Please upload at least one product photo.")
            return

        try:
            with st.spinner("Reading price list and matching photos..."):
                price_df = load_price_dataframe(price_file)
                matched_df, temp_dir = match_photos_to_prices(photo_files, price_df)

            st.success(f"Matched {len(matched_df)} photos to the price list.")

            # Show a small preview
            st.subheader("Preview of matches")
            st.dataframe(
                matched_df[["PHOTO_FILE", "CODE", "DESCRIPTION", "PRICE_A_INCL"]].head(20),
                use_container_width=True
            )

            with st.spinner("Building Excel with thumbnails..."):
                excel_bytes = build_excel_with_thumbnails(matched_df, temp_dir)

            st.download_button(
                "Download Excel catalogue with photos",
                data=excel_bytes,
                file_name="photo_catalogue.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

        except Exception as e:
            st.error(f"Something went wrong: {e}")


if __name__ == "__main__":
    main()
