import streamlit as st
import pandas as pd
import os
import re
from PIL import Image
from io import BytesIO
from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Proto Catalogue Builder", layout="wide")

st.title("📸 Proto Catalogue Builder")
st.write("Upload your Excel + photo folder and I'll match everything automatically.")

# ---------------------------
# 1. HELPER: Identify columns
# ---------------------------

def find_column(columns, patterns):
    """Find best matching column name."""
    for col in columns:
        for pat in patterns:
            if re.search(pat, col.lower()):
                return col
    return None


# -------------------------------------
# 2. Extract code from image filenames
# -------------------------------------

def extract_code_from_filename(filename):
    match = re.search(r"(\d{6,15})", filename)
    return match.group(1) if match else None


# -----------------------------------------------------
# 3. Match photos to Excel (smart matching)
# -----------------------------------------------------

def match_photos_to_excel(df, image_folder):

    images = [f for f in os.listdir(image_folder)
              if f.lower().endswith((".jpg", ".jpeg", ".png"))]

    rows = []

    for img_file in images:
        photo_code = extract_code_from_filename(img_file)
        if not photo_code:
            continue

        match = df[df["CODE"].astype(str).str.contains(photo_code[:8])]

        if match.empty:
            match = df[df["CODE"].astype(str).str.startswith(photo_code[:6])]

        if not match.empty:
            row = match.iloc[0]
            rows.append({
                "FILENAME": img_file,
                "CODE": row["CODE"],
                "DESCRIPTION": row["DESCRIPTION"],
                "PRICE_INCL": row["PRICE_INCL"]
            })

    return pd.DataFrame(rows)


# ---------------------------------------------------------
# 4. BUILD EXCEL WITH THUMBNAILS
# ---------------------------------------------------------

def build_excel_with_thumbnails(df, image_folder):

    wb = Workbook()
    ws = wb.active
    ws.title = "Catalogue"

    headers = ["PHOTO", "CODE", "DESCRIPTION", "PRICE_INCL", "FILENAME"]
    ws.append(headers)

    row_index = 2

    for _, row in df.iterrows():
        img_path = os.path.join(image_folder, row["FILENAME"])

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

        ws[f"B{row_index}"] = row["CODE"]
        ws[f"C{row_index}"] = row["DESCRIPTION"]
        ws[f"D{row_index}"] = row["PRICE_INCL"]
        ws[f"E{row_index}"] = row["FILENAME"]

        ws.row_dimensions[row_index].height = 120
        row_index += 1

    ws.column_dimensions["A"].width = 25
    ws.column_dimensions["B"].width = 18
    ws.column_dimensions["C"].width = 45
    ws.column_dimensions["D"].width = 18
    ws.column_dimensions["E"].width = 30

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output


# -------------------------------------------------------------
# 5. STREAMLIT UI
# -------------------------------------------------------------

uploaded_excel = st.file_uploader("Upload Excel Price File", type=["xlsx", "xls"])
uploaded_folder = st.text_input("Enter path to your photo folder (e.g. C:/Users/Photos)")

if uploaded_excel and uploaded_folder:

    with st.spinner("Reading Excel file…"):

        df = pd.read_excel(uploaded_excel)
        cols = df.columns

        col_code = find_column(cols, [r"code", r"item", r"product"])
        col_desc = find_column(cols, [r"desc", r"description", r"details"])
        col_price = find_column(cols, [r"incl", r"price a", r"price_incl", r"vat incl"])

        if not all([col_code, col_desc, col_price]):
            st.error("Could not automatically detect CODE, DESCRIPTION or PRICE columns.")
        else:
            df = df.rename(columns={
                col_code: "CODE",
                col_desc: "DESCRIPTION",
                col_price: "PRICE_INCL"
            })

            st.success("Excel columns detected successfully!")

            with st.spinner("Matching photos…"):
                matched_df = match_photos_to_excel(df, uploaded_folder)

            st.write("### ✅ Matched Items")
            st.dataframe(matched_df)

            if not matched_df.empty:
                excel_file = build_excel_with_thumbnails(matched_df, uploaded_folder)

                st.download_button(
                    "📥 Download Catalogue Excel",
                    data=excel_file,
                    file_name="catalogue.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
