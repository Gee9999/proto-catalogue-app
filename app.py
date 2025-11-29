import streamlit as st
import pandas as pd
import tempfile
import os
import re
from PIL import Image
from fpdf import FPDF
from io import BytesIO

st.set_page_config(page_title="Photo Catalogue Builder", layout="wide")


# ---------------------------------------------
# EXTRACT NUMERIC CODE FROM PHOTO FILENAME
# ---------------------------------------------
def extract_code_from_filename(filename):
    m = re.search(r"(\d{6,12})", filename)
    return m.group(1) if m else None


# ---------------------------------------------
# CLEAN + DETECT IMPORTANT COLUMNS
# ---------------------------------------------
def normalize(col):
    return col.lower().replace(" ", "").replace("-", "").replace("_", "").replace("\n", "")


def find_column(df, keywords):
    for col in df.columns:
        clean = normalize(col)
        for k in keywords:
            if k in clean:
                return col
    return None


# ---------------------------------------------
# MATCH PHOTOS TO EXCEL DATA
# ---------------------------------------------
def match_photos_to_prices(photo_files, price_df):
    # detect columns
    code_col = find_column(price_df, ["code", "itemcode", "barcode"])
    desc_col = find_column(price_df, ["description", "desc"])
    price_col = find_column(price_df, ["price", "aincl", "incl", "vat"])

    if not code_col or not desc_col or not price_col:
        raise ValueError(f"""
        ❌ Could not detect CODE / DESCRIPTION / PRICE column.
        Columns found: {list(price_df.columns)}
        """)

    # Normalise code column to string
    price_df["CODE_STR"] = price_df[code_col].astype(str).str.replace(".0", "", regex=False)

    # Build dictionary for fast lookup
    lookup = {
        row["CODE_STR"]: {
            "DESCRIPTION": row[desc_col],
            "PRICE": row[price_col]
        }
        for _, row in price_df.iterrows()
    }

    results = []

    for file in photo_files:
        filename = file.name
        code = extract_code_from_filename(filename)

        if not code:
            results.append({
                "PHOTO_FILE": filename,
                "CODE": "",
                "DESCRIPTION": "",
                "PRICE": ""
            })
            continue

        best_match = lookup.get(code, None)

        if best_match:
            desc = best_match["DESCRIPTION"]
            price = best_match["PRICE"]
        else:
            desc = ""
            price = ""

        results.append({
            "PHOTO_FILE": filename,
            "CODE": code,
            "DESCRIPTION": desc,
            "PRICE": price,
        })

    return pd.DataFrame(results)


# ---------------------------------------------
# BUILD EXCEL WITH THUMBNAILS
# ---------------------------------------------
def build_excel_with_images(df, temp_dir):
    from openpyxl import Workbook
    from openpyxl.drawing.image import Image as XLImage

    wb = Workbook()
    ws = wb.active
    ws.title = "Catalogue"

    ws.append(["Photo", "Code", "Description", "Price-A Incl", "Filename"])

    for index, row in df.iterrows():
        photo_path = os.path.join(temp_dir, row["PHOTO_FILE"])

        # Insert row
        excel_row = index + 2  # header is row 1
        ws.cell(row=excel_row, column=2, value=row["CODE"])
        ws.cell(row=excel_row, column=3, value=row["DESCRIPTION"])
        ws.cell(row=excel_row, column=4, value=row["PRICE"])
        ws.cell(row=excel_row, column=5, value=row["PHOTO_FILE"])

        # Insert thumbnail
        try:
            img = Image.open(photo_path)
            img.thumbnail((100, 100))
            thumb_path = os.path.join(temp_dir, f"thumb_{row['PHOTO_FILE']}.jpg")
            img.save(thumb_path)

            xl_img = XLImage(thumb_path)
            ws.add_image(xl_img, f"A{excel_row}")

        except Exception:
            pass

    out = BytesIO()
    wb.save(out)
    return out.getvalue()


# ---------------------------------------------
# BUILD PDF — TEXT ALWAYS UNDER THE PHOTO
# ---------------------------------------------
def build_pdf(df, temp_dir, cell_w=63, cell_h=63):
    pdf = FPDF(unit="mm", format="A4")
    pdf.add_page()
    pdf.set_auto_page_break(auto=True, margin=10)

    x_start = 10
    y_start = 20
    x = x_start
    y = y_start
    count = 0

    for _, row in df.iterrows():
        photo_path = os.path.join(temp_dir, row["PHOTO_FILE"])

        try:
            img = Image.open(photo_path)
            img.thumbnail((cell_w, cell_h))
            photo_tmp = os.path.join(temp_dir, f"pdf_{row['PHOTO_FILE']}.jpg")
            img.save(photo_tmp)

            # Insert image
            pdf.image(photo_tmp, x=x, y=y, w=cell_w, h=cell_h)

        except Exception:
            pass

        # Insert text UNDER the image
        pdf.set_xy(x, y + cell_h + 2)
        pdf.set_font("Arial", size=8)

        pdf.multi_cell(cell_w, 4, f"{row['CODE']}", 0, "L")
        pdf.multi_cell(cell_w, 4, f"{row['DESCRIPTION']}", 0, "L")
        pdf.multi_cell(cell_w, 4, f"Price: {row['PRICE']}", 0, "L")

        # Next column
        x += cell_w + 10
        count += 1

        # New row every 3 columns
        if count % 3 == 0:
            x = x_start
            y += cell_h + 25

        # New page if needed
        if y > 250:
            pdf.add_page()
            x = x_start
            y = y_start

    return pdf.output(dest="S").encode("latin1")


# ---------------------------------------------
# MAIN STREAMLIT APP
# ---------------------------------------------
def main():
    st.title("📸 Photo Catalogue Builder")

    st.write("""
    Upload:
    **1) Price Excel (any layout — must contain CODE, DESCRIPTION & VAT inclusive price)**  
    **2) Product photos (filename must include the code)**  
    """)

    price_file = st.file_uploader("Upload price Excel", type=["xls", "xlsx"])
    photo_files = st.file_uploader("Upload product photos", type=["jpg", "jpeg", "png"], accept_multiple_files=True)

    # Photo size option
    size_choice = st.selectbox("Choose photo size for PDF:", [
        "Large (120x120)",
        "Medium (90x90)",
        "Small (60x60)"
    ])

    if size_choice == "Large (120x120)":
        cell_w = cell_h = 120
    elif size_choice == "Medium (90x90)":
        cell_w = cell_h = 90
    else:
        cell_w = cell_h = 60

    if st.button("Generate catalogue"):
        if not price_file or not photo_files:
            st.error("Please upload BOTH the price Excel and product photos.")
            return

        with tempfile.TemporaryDirectory() as temp_dir:

            # Save photos
            for f in photo_files:
                with open(os.path.join(temp_dir, f.name), "wb") as out:
                    out.write(f.read())

            # Load Excel
            df = pd.read_excel(price_file)

            # Match
            matched_df = match_photos_to_prices(photo_files, df)

            # Build Excel
            excel_bytes = build_excel_with_images(matched_df, temp_dir)

            # Build PDF
            pdf_bytes = build_pdf(matched_df, temp_dir, cell_w=cell_w, cell_h=cell_h)

            # Download links
            st.success("Done! Download your files below:")

            st.download_button("📥 Download Excel", excel_bytes, file_name="photo_catalogue.xlsx")
            st.download_button("📥 Download PDF", pdf_bytes, file_name="photo_catalogue.pdf")


main()
