import streamlit as st
import pandas as pd
import os
import tempfile
from fpdf import FPDF
from PIL import Image
import io
import re

st.set_page_config(page_title="Photo Catalogue Builder", layout="wide")

# ------------------------------------------
# AUTO-DETECT COLUMNS IN PRICE EXCEL
# ------------------------------------------
def normalize_column(col: str):
    col = col.strip().lower().replace(" ", "").replace("-", "")
    return col

def detect_columns(df):
    df_cols = {normalize_column(c): c for c in df.columns}

    code_col = None
    desc_col = None
    price_col = None

    for key, real in df_cols.items():
        if key.startswith("code"):
            code_col = real
        if key.startswith("desc"):
            desc_col = real
        if key.startswith("priceaincl") or key.startswith("priceincl") or key.startswith("pricea"):
            price_col = real

    if not code_col or not desc_col or not price_col:
        raise ValueError("Excel must contain columns: CODE, DESCRIPTION, PRICE-A INCL.")

    return code_col, desc_col, price_col

# ------------------------------------------
# MATCH PHOTOS TO PRICES (PREFIX MATCH)
# ------------------------------------------
def extract_prefix(filename):
    """
    Extract longest leading numeric prefix from filename.
    Example: "8613900001-25pcs.jpg" → "8613900001"
    """
    m = re.match(r"^(\d+)", filename)
    return m.group(1) if m else None

def match_photos_to_prices(photo_files, price_df):
    price_df["CODE_STR"] = price_df["CODE"].astype(str)

    # Make dictionary for fast lookup
    lookup = price_df.set_index("CODE_STR")[["CODE", "DESCRIPTION", "PRICE_A_INCL"]].to_dict("index")

    rows = []
    for file in photo_files:
        fname = file.name
        prefix = extract_prefix(fname)

        if prefix and prefix in lookup:
            data = lookup[prefix]
            rows.append({
                "PHOTO_FILE": fname,
                "CODE": data["CODE"],
                "DESCRIPTION": data["DESCRIPTION"],
                "PRICE_A_INCL": data["PRICE_A_INCL"],
                "FILE_OBJ": file
            })
        else:
            rows.append({
                "PHOTO_FILE": fname,
                "CODE": prefix if prefix else "",
                "DESCRIPTION": "",
                "PRICE_A_INCL": "",
                "FILE_OBJ": file
            })

    return pd.DataFrame(rows)

# ------------------------------------------
# BUILD PDF 3×3 WITH TEXT UNDER IMAGES
# ------------------------------------------
def build_pdf(df, temp_dir):

    # Save temp images
    for idx, item in df.iterrows():
        im = Image.open(item["FILE_OBJ"])
        out_path = os.path.join(temp_dir, f"img_{idx}.jpg")
        im.save(out_path)
        df.at[idx, "TEMP_PATH"] = out_path

    pdf = FPDF("P", "mm", "A4")
    pdf.set_auto_page_break(auto=True, margin=10)

    pdf.add_page()
    pdf.set_font("Arial", "B", 14)
    pdf.cell(0, 10, "Photo Catalogue", ln=1, align="C")

    cell_w = 190 / 3
    cell_h = 70
    fixed_image_height = 40

    x_start = 10
    y_start = 25

    col = 0
    row = 0

    for idx, item in df.iterrows():
        x = x_start + col * cell_w
        y = y_start + row * cell_h

        img_path = item["TEMP_PATH"]
        if os.path.exists(img_path):
            try:
                pdf.image(
                    img_path,
                    x=x + 2,
                    y=y,
                    w=cell_w - 4,
                    h=fixed_image_height,
                    keep_aspect_ratio=True
                )
            except:
                pass

        # TEXT under image
        pdf.set_xy(x + 2, y + fixed_image_height + 2)
        pdf.set_font("Arial", size=8)

        code_line = f"Code: {item['CODE']}"
        desc_line = str(item["DESCRIPTION"])[:60]
        price_line = (
            f"Price: {item['PRICE_A_INCL']:,.2f}"
            if isinstance(item["PRICE_A_INCL"], (float, int))
            else "Price: -"
        )

        text_w = cell_w - 4

        pdf.multi_cell(text_w, 4, code_line)
        pdf.set_x(x + 2)
        pdf.multi_cell(text_w, 4, desc_line)
        pdf.set_x(x + 2)
        pdf.multi_cell(text_w, 4, price_line)

        col += 1
        if col == 3:
            col = 0
            row += 1

        if row == 3:
            pdf.add_page()
            row = 0

    # FIX for bytearray output
    raw = pdf.output(dest="S")
    return raw if isinstance(raw, (bytes, bytearray)) else raw.encode("latin1")

# ------------------------------------------
# BUILD EXCEL OUTPUT
# ------------------------------------------
def build_excel(df):
    output = io.BytesIO()
    df.to_excel(output, index=False)
    return output.getvalue()

# ------------------------------------------
# STREAMLIT UI
# ------------------------------------------
def main():
    st.title("Photo Catalogue Builder")

    st.write("Upload your price Excel file (must contain CODE, DESCRIPTION, PRICE-A INCL).")
    price_file = st.file_uploader("Upload price Excel", type=["xls", "xlsx"])

    st.write("Upload product photos (filenames must contain the numeric code).")
    photo_files = st.file_uploader("Upload product photos", type=["jpg", "jpeg", "png"], accept_multiple_files=True)

    if price_file and photo_files:
        if st.button("Generate Catalogue"):
            try:
                price_df_raw = pd.read_excel(price_file)
                code_col, desc_col, price_col = detect_columns(price_df_raw)

                price_df = price_df_raw[[code_col, desc_col, price_col]].copy()
                price_df.columns = ["CODE", "DESCRIPTION", "PRICE_A_INCL"]

                matched_df = match_photos_to_prices(photo_files, price_df)

                with tempfile.TemporaryDirectory() as tmp:
                    pdf_bytes = build_pdf(matched_df, tmp)
                    excel_bytes = build_excel(matched_df)

                st.success("Catalogue generated successfully!")

                st.download_button("Download PDF Catalogue", pdf_bytes, file_name="catalogue.pdf", mime="application/pdf")
                st.download_button("Download Excel File", excel_bytes, file_name="catalogue.xlsx", mime="application/vnd.ms-excel")

            except Exception as e:
                st.error(f"Something went wrong: {e}")

main()
