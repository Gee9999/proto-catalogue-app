import streamlit as st
import pandas as pd
import tempfile
import os
from PIL import Image
from fpdf import FPDF

# ----------------------------------------
# 1) SMART PRICE FILE LOADER
# ----------------------------------------
def load_price_file(file):
    df = pd.read_excel(file, dtype=str)

    # Clean column names
    clean_cols = {
        col: ''.join(c for c in col.upper().replace(" ", "").replace(".", "").replace("-", "").replace("_", ""))
        for col in df.columns
    }
    df.rename(columns=clean_cols, inplace=True)

    code_col = None
    desc_col = None
    price_col = None

    for col in df.columns:
        c = col.upper()
        if "CODE" in c:
            code_col = col
        if "DESC" in c or "DESCRIPTION" in c:
            desc_col = col
        if "PRICE" in c and "A" in c and ("INCL" in c or "VAT" in c):
            price_col = col

    if code_col is None:
        raise ValueError("Could not find CODE column.")
    if desc_col is None:
        raise ValueError("Could not find DESCRIPTION column.")
    if price_col is None:
        raise ValueError("Could not find PRICE-A INCL column.")

    out = df[[code_col, desc_col, price_col]].copy()
    out.columns = ["CODE", "DESCRIPTION", "PRICE_A_INCL"]

    # Clean CODE
    out["CODE"] = out["CODE"].astype(str).str.replace(".0", "", regex=False)

    return out


# ----------------------------------------
# 2) EXTRACT CODE FROM PHOTO FILENAME
# ----------------------------------------
def extract_code_from_filename(filename):
    base = os.path.splitext(filename)[0]
    digits = ''.join([c for c in base if c.isdigit()])
    return digits[:10]  # first 10 digits is your code


# ----------------------------------------
# 3) MATCH PHOTOS TO PRICE DATA
# ----------------------------------------
def match_photos_to_prices(photo_files, price_df):
    price_df["CODE_STR"] = price_df["CODE"].astype(str)
    price_dict = price_df.set_index("CODE_STR")[["DESCRIPTION", "PRICE_A_INCL"]].to_dict("index")

    rows = []
    for photo in photo_files:
        code = extract_code_from_filename(photo.name)
        desc = ""
        price = ""
        if code in price_dict:
            desc = price_dict[code]["DESCRIPTION"]
            price = price_dict[code]["PRICE_A_INCL"]

        rows.append({
            "PHOTO_FILE": photo,
            "CODE": code,
            "DESCRIPTION": desc,
            "PRICE_A_INCL": price
        })

    return pd.DataFrame(rows)


# ----------------------------------------
# 4) BUILD PDF WITH 120x120 IMAGES
# ----------------------------------------
def build_pdf(df, temp_dir, thumb_size=120):
    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=10)
    pdf.add_page()

    pdf.set_font("Arial", "B", 16)
    pdf.cell(0, 10, "Photo Catalogue", ln=1, align="C")

    cell_w = 70
    cell_h = 70
    items_per_row = 3

    col = 0

    for _, row in df.iterrows():
        photo_file = row["PHOTO_FILE"]
        code = row["CODE"]
        desc = row["DESCRIPTION"]
        price = row["PRICE_A_INCL"]

        img = Image.open(photo_file)
        img.thumbnail((thumb_size, thumb_size))
        temp_img_path = os.path.join(temp_dir, f"thumb_{code}.jpg")
        img.save(temp_img_path, "JPEG")

        if col == 0:
            pdf.ln(10)

        x = pdf.get_x()
        y = pdf.get_y()

        pdf.image(temp_img_path, x=x, y=y, w=cell_w, h=cell_h)

        pdf.set_xy(x, y + cell_h + 2)
        pdf.set_font("Arial", size=8)

        pdf.multi_cell(cell_w, 4, f"{code}\n{desc}\n{price}", border=0, align="L")

        col += 1
        if col >= items_per_row:
            col = 0
            pdf.ln(cell_h + 20)
        else:
            pdf.set_xy(x + cell_w + 5, y)

    return pdf.output(dest="S").encode("latin1")


# ----------------------------------------
# 5) BUILD EXCEL WITH THUMBNAILS (120×120)
# ----------------------------------------
def build_excel(df, temp_dir, thumb_size=120):
    from openpyxl import Workbook
    from openpyxl.drawing.image import Image as XLImage

    wb = Workbook()
    ws = wb.active
    ws.title = "Catalogue"

    ws.append(["PHOTO", "CODE", "DESCRIPTION", "PRICE_A_INCL"])

    row_idx = 2
    for _, row in df.iterrows():
        photo_file = row["PHOTO_FILE"]
        code = row["CODE"]
        desc = row["DESCRIPTION"]
        price = row["PRICE_A_INCL"]

        img = Image.open(photo_file)
        img.thumbnail((thumb_size, thumb_size))

        temp_img = os.path.join(temp_dir, f"excel_{code}.jpg")
        img.save(temp_img)

        xl_img = XLImage(temp_img)
        xl_img.width = thumb_size
        xl_img.height = thumb_size

        ws.row_dimensions[row_idx].height = 100
        ws.add_image(xl_img, f"A{row_idx}")

        ws[f"B{row_idx}"] = code
        ws[f"C{row_idx}"] = desc
        ws[f"D{row_idx}"] = price

        row_idx += 1

    excel_bytes = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    wb.save(excel_bytes.name)
    excel_bytes.seek(0)
    return excel_bytes.read()


# ----------------------------------------
# 6) STREAMLIT APP UI
# ----------------------------------------
def main():
    st.title("📸 Photo Catalogue Builder")

    price_file = st.file_uploader("Upload price Excel file", type=["xlsx", "xls"])
    photo_files = st.file_uploader("Upload product photos", type=["jpg", "jpeg", "png"], accept_multiple_files=True)

    if st.button("Generate Catalogue"):
        if not price_file or not photo_files:
            st.error("Please upload both Excel and photos.")
            return

        with st.spinner("Processing..."):
            temp_dir = tempfile.mkdtemp()

            df_price = load_price_file(price_file)
            df_matched = match_photos_to_prices(photo_files, df_price)

            pdf_bytes = build_pdf(df_matched, temp_dir, thumb_size=120)
            excel_bytes = build_excel(df_matched, temp_dir, thumb_size=120)

        st.success("Done!")

        st.download_button("Download PDF", data=pdf_bytes, file_name="catalogue.pdf", mime="application/pdf")
        st.download_button("Download Excel", data=excel_bytes, file_name="catalogue.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


if __name__ == "__main__":
    main()
