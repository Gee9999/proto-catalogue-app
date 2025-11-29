import streamlit as st
import pandas as pd
import re
import os
import tempfile
import io
from PIL import Image as PILImage
from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage

st.set_page_config(page_title="Photo Catalogue Builder", layout="wide")


# ---------------------------------------------------------
# 1. Extract numeric code from filename
# ---------------------------------------------------------
def extract_code_from_filename(filename: str) -> str:
    """
    Extracts the first long numeric sequence from a filename.
    e.g. "8613900001-25pcs.jpg" → "8613900001"
    """
    match = re.search(r"(\d{6,})", filename)
    return match.group(1) if match else ""


# ---------------------------------------------------------
# 2. Load Excel and auto-find CODE / DESCRIPTION / PRICE columns
# ---------------------------------------------------------
def load_price_dataframe(uploaded_file):
    df = pd.read_excel(uploaded_file)

    # Normalize column names
    clean_cols = {col: col.strip().lower().replace(" ", "").replace("-", "") for col in df.columns}
    df.rename(columns=clean_cols, inplace=True)

    # Identify CODE column
    code_col = next((col for col in df.columns if "code" in col), None)
    if not code_col:
        raise ValueError("Could not find CODE column in Excel.")

    # Identify DESCRIPTION column
    desc_col = next((col for col in df.columns if "desc" in col or "description" in col), None)
    if not desc_col:
        raise ValueError("Could not find DESCRIPTION column in Excel.")

    # Identify PRICE column
    price_col = None
    for col in df.columns:
        if "price" in col or "incl" in col or "aincl" in col or "vat" in col:
            price_col = col
            break

    if not price_col:
        raise ValueError("Could not find PRICE-A INCL column.")

    # Standardise
    df = df[[code_col, desc_col, price_col]].copy()
    df.columns = ["CODE", "DESCRIPTION", "PRICE_A_INCL"]

    # Clean CODE to string (important)
    df["CODE"] = df["CODE"].astype(str).str.extract(r"(\d{6,})")

    return df


# ---------------------------------------------------------
# 3. Match photos to price list
# ---------------------------------------------------------
def match_photos_to_prices(photo_files, price_df):
    # Extract all excel codes into strings
    excel_codes = price_df["CODE"].astype(str).tolist()

    matched_rows = []

    temp_dir = tempfile.mkdtemp()

    for uploaded in photo_files:
        filename = uploaded.name
        photo_code = extract_code_from_filename(filename)

        # Save temp image
        temp_path = os.path.join(temp_dir, filename)
        with open(temp_path, "wb") as f:
            f.write(uploaded.getvalue())

        # Find exact match
        if photo_code in excel_codes:
            row = price_df.loc[price_df["CODE"] == photo_code].iloc[0]
            description = row["DESCRIPTION"]
            price = row["PRICE_A_INCL"]
        else:
            description = ""
            price = ""

        matched_rows.append(dict(
            PHOTO_FILE=filename,
            CODE_FROM_PHOTO=photo_code,
            CODE=photo_code if photo_code in excel_codes else "",
            DESCRIPTION=description,
            PRICE_A_INCL=price
        ))

    matched_df = pd.DataFrame(matched_rows)
    return matched_df, temp_dir


# ---------------------------------------------------------
# 4. Build Excel with thumbnails
# ---------------------------------------------------------
def build_excel_with_thumbnails(matched_df: pd.DataFrame, temp_dir: str, thumb_size=(60, 60)) -> io.BytesIO:
    wb = Workbook()
    ws = wb.active
    ws.title = "Catalogue"

    ws.column_dimensions["A"].width = 18 if max(thumb_size) > 80 else 14
    ws.column_dimensions["B"].width = 14
    ws.column_dimensions["C"].width = 45
    ws.column_dimensions["D"].width = 14

    ws.append(["PHOTO", "CODE", "DESCRIPTION", "PRICE-A INCL"])

    for idx, row in matched_df.iterrows():
        excel_row = idx + 2

        ws.row_dimensions[excel_row].height = max(thumb_size) + 10

        code = row.get("CODE", "")
        desc = row.get("DESCRIPTION", "")
        price = row.get("PRICE_A_INCL", "")

        ws.cell(row=excel_row, column=2, value=code)
        ws.cell(row=excel_row, column=3, value=desc)
        ws.cell(row=excel_row, column=4, value=price)

        img_path = os.path.join(temp_dir, row["PHOTO_FILE"])
        if os.path.exists(img_path):
            try:
                pil_img = PILImage.open(img_path)
                pil_img.thumbnail(thumb_size)

                bio = io.BytesIO()
                pil_img.save(bio, format="PNG")
                bio.seek(0)
                bio.name = "thumb.png"

                excel_img = XLImage(bio)
                excel_img.anchor = f"A{excel_row}"
                ws.add_image(excel_img)
            except:
                pass

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out


# ---------------------------------------------------------
# 5. UI
# ---------------------------------------------------------
def main():
    st.title("📸 Photo Catalogue Builder")

    st.write("""
    Upload your **price list Excel** and **product photos**, and I’ll generate:
    - A **perfectly matched Excel** catalogue with thumbnails  
    - CODE, DESCRIPTION, PRICE underneath each image  
    """)

    st.subheader("1️⃣ Upload Price Excel (any format)")
    price_file = st.file_uploader("Upload Excel", type=["xls", "xlsx"])

    st.subheader("2️⃣ Upload Product Photos")
    photo_files = st.file_uploader("Upload JPG / PNG photos", type=["jpg", "jpeg", "png"], accept_multiple_files=True)

    size_choice = st.selectbox(
        "Choose image thumbnail size:",
        ["Auto (recommended)", "Small (60x60)", "Medium (90x90)", "Large (120x120)"]
    )

    if st.button("Generate catalogue"):
        try:
            if not price_file:
                st.error("Please upload a price Excel file first.")
                return

            if not photo_files:
                st.error("Please upload at least one product photo.")
                return

            price_df = load_price_dataframe(price_file)

            if size_choice == "Small (60x60)":
                thumb_size = (60, 60)
            elif size_choice == "Medium (90x90)":
                thumb_size = (90, 90)
            elif size_choice == "Large (120x120)":
                thumb_size = (120, 120)
            else:
                n = len(photo_files)
                thumb_size = (120, 120) if n <= 150 else (60, 60)

            matched_df, temp_dir = match_photos_to_prices(photo_files, price_df)

            excel_bytes = build_excel_with_thumbnails(matched_df, temp_dir, thumb_size)

            st.success("Catalogue generated successfully!")

            st.download_button(
                "Download Excel Catalogue",
                data=excel_bytes,
                file_name="photo_catalogue.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

            st.subheader("Preview (first 20 rows)")
            st.dataframe(matched_df.head(20))

        except Exception as e:
            st.error(f"Something went wrong: {e}")


if __name__ == "__main__":
    main()
