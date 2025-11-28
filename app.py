import streamlit as st
import pandas as pd
import os
import tempfile
from PIL import Image
from fpdf import FPDF
from io import BytesIO

st.set_page_config(page_title="Proto Catalogue Builder", layout="wide")

# -----------------------------
# 1) Extract code from filename
# -----------------------------
def extract_code_from_filename(filename):
    code = ""
    for ch in filename:
        if ch.isdigit():
            code += ch
        else:
            break
    return code


# ------------------------------------
# 2) Build Excel with thumbnails
# ------------------------------------
def build_excel_with_thumbnails(df, temp_dir):
    from openpyxl import Workbook
    from openpyxl.drawing.image import Image as XLImage

    wb = Workbook()
    ws = wb.active
    ws.append(["Thumbnail", "Code", "Description", "Price Incl", "Filename"])

    ws.column_dimensions["A"].width = 22
    ws.column_dimensions["B"].width = 15
    ws.column_dimensions["C"].width = 50
    ws.column_dimensions["D"].width = 15
    ws.column_dimensions["E"].width = 40

    for idx, row in df.iterrows():
        img_path = row["img_path"]

        thumb_path = os.path.join(temp_dir, f"thumb_{idx}.jpg")
        try:
            img = Image.open(img_path)
            img.thumbnail((160, 160))
            img.save(thumb_path)
        except:
            continue

        ws.append(["", row["code"], row["description"], row["price_incl"], row["filename"]])

        xl_img = XLImage(thumb_path)
        xl_img.width = 160
        xl_img.height = 160
        ws.add_image(xl_img, f"A{ws.max_row}")

    out = BytesIO()
    wb.save(out)
    out.seek(0)
    return out


# ------------------------------------
# 3) Build PDF (3 per row)
# ------------------------------------
def build_pdf(df):
    pdf = FPDF("P", "mm", "A4")
    pdf.set_auto_page_break(auto=True, margin=8)

    cell_w = 60
    cell_h = 60

    for i in range(0, len(df), 3):
        row = df.iloc[i:i+3]
        pdf.add_page()

        # FIRST ROW: IMAGES
        x = 10
        for _, r in row.iterrows():
            try:
                img = Image.open(r["img_path"])
                img_path = r["img_path"]
                pdf.image(img_path, x=x, y=20, w=cell_w)
            except:
                pass
            x += cell_w + 10

        # SECOND ROW: DESCRIPTION + PRICE + CODE
        x = 10
        y_text = 85
        pdf.set_font("Arial", size=10)

        for _, r in row.iterrows():
            pdf.set_xy(x, y_text)
            pdf.multi_cell(cell_w, 5, f"{r['description']}", 0, "L")

            pdf.set_xy(x, y_text + 18)
            pdf.set_font("Arial", size=10)
            pdf.cell(cell_w, 5, f"Price: {r['price_incl']}", 0, 2)

            pdf.set_font("Arial", size=9)
            pdf.cell(cell_w, 5, f"Code: {r['code']}", 0, 2)

            x += cell_w + 10

    out = BytesIO()
    pdf.output(out)
    out.seek(0)
    return out


# ------------------------------------
# STREAMLIT UI
# ------------------------------------
st.title("📸 Proto Catalogue Builder")
st.subheader("Auto-match photos → price + description → Excel + PDF")


uploaded_excel = st.file_uploader("Upload price Excel file", type=["xlsx"])
uploaded_images = st.file_uploader("Upload product photos", type=["jpg", "jpeg", "png"], accept_multiple_files=True)

if uploaded_excel and uploaded_images:
    st.success("Files uploaded successfully!")

    df = pd.read_excel(uploaded_excel)

    # Standardize column names
    df.columns = df.columns.str.strip().str.upper()

    if not {"CODE", "DESCRIPTION", "PRICE-AINCL."}.issubset(df.columns):
        st.error("Excel must contain columns: CODE, DESCRIPTION, PRICE-AINCL.")
        st.stop()

    price_map = {}
    for _, r in df.iterrows():
        code = str(r["CODE"]).split(".")[0]
        price = r["PRICE-AINCL."]
        desc = str(r["DESCRIPTION"])
        price_map[code] = (desc, price)

    matched_rows = []
    temp_dir = tempfile.mkdtemp()

    for img in uploaded_images:
        filename = img.name
        file_path = os.path.join(temp_dir, filename)

        with open(file_path, "wb") as f:
            f.write(img.read())

        code = extract_code_from_filename(filename)
        desc, price = price_map.get(code, ("", ""))

        matched_rows.append({
            "filename": filename,
            "code": code,
            "description": desc,
            "price_incl": price,
            "img_path": file_path
        })

    matched_df = pd.DataFrame(matched_rows)

    st.subheader("Download Excel")
    excel_file = build_excel_with_thumbnails(matched_df, temp_dir)
    st.download_button("Download Excel", excel_file, file_name="catalogue.xlsx")

    st.subheader("Download PDF")
    pdf_file = build_pdf(matched_df)
    st.download_button("Download PDF", pdf_file, file_name="catalogue.pdf")

