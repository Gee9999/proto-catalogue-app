import streamlit as st
import pandas as pd
import tempfile
import os
from fpdf import FPDF
from PIL import Image
from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage

st.set_page_config(page_title="Photo Catalogue Builder", layout="wide")

# -----------------------------
# Extract numeric code from filename
# -----------------------------
def extract_code_from_filename(filename):
    digits = "".join(c for c in filename if c.isdigit())
    return digits[:10]  # Your codes are always first 10 digits

# -----------------------------
# Find closest code match (exact or nearest)
# -----------------------------
def get_best_match(code_str, code_dict):
    if code_str in code_dict:
        return code_str
    
    # fallback – closest numeric distance
    try:
        target = int(code_str)
        best = None
        best_diff = 999999999
        for k in code_dict:
            diff = abs(int(k) - target)
            if diff < best_diff:
                best_diff = diff
                best = k
        return best
    except:
        return None

# -----------------------------
# Match photos with price Excel
# -----------------------------
def match_photos_to_prices(photo_files, price_df):
    price_df["CODE_STR"] = price_df["CODE"].astype(str).str.extract(r"(\d+)")
    price_df["CODE_STR"] = price_df["CODE_STR"].str[:10]

    # Make index unique by grouping duplicates and taking first
    price_df = price_df.groupby("CODE_STR", as_index=False).first()

    code_dict = price_df.set_index("CODE_STR")[["CODE", "DESCRIPTION", "PRICE_A_INCL"]].to_dict("index")

    results = []

    for file in photo_files:
        filename = file.name
        extracted = extract_code_from_filename(filename)
        best = get_best_match(extracted, code_dict)

        if best:
            row = code_dict[best]
            results.append({
                "PHOTO_FILE": filename,
                "CODE": row["CODE"],
                "DESCRIPTION": row["DESCRIPTION"],
                "PRICE_A_INCL": row["PRICE_A_INCL"]
            })
        else:
            results.append({
                "PHOTO_FILE": filename,
                "CODE": "NOT FOUND",
                "DESCRIPTION": "",
                "PRICE_A_INCL": ""
            })

    return pd.DataFrame(results)

# -----------------------------
# Build PDF (already perfect)
# -----------------------------
def build_pdf(df, temp_dir):
    pdf = FPDF("P", "mm", "A4")
    pdf.add_page()
    pdf.set_font("Arial", "B", 16)
    pdf.cell(0, 10, "Photo Catalogue", ln=1, align="C")

    pdf.set_font("Arial", size=10)

    cell_w = 90
    cell_h = 70
    x_start = 10
    y_start = 30

    col = 0
    row_y = y_start

    for i, r in df.iterrows():
        if col == 2:
            col = 0
            row_y += cell_h + 20
        if row_y > 260:
            pdf.add_page()
            row_y = y_start
            col = 0

        x = x_start + col * (cell_w + 10)

        # Save image to temp
        img_path = os.path.join(temp_dir, f"img_{i}.jpg")
        with open(img_path, "wb") as f:
            f.write(r["FILE_BYTES"])

        # Photo
        pdf.image(img_path, x=x, y=row_y, w=60)

        # Text ALWAYS below image
        pdf.set_xy(x, row_y + 62)
        pdf.multi_cell(60, 5, f"{r['CODE']}\n{r['DESCRIPTION']}\nR {r['PRICE_A_INCL']}", align="L")

        col += 1

    return pdf.output(dest="S").encode("latin1")

# -----------------------------
# Build Excel with thumbnails
# -----------------------------
def build_excel_with_thumbs(df, temp_dir):
    wb = Workbook()
    ws = wb.active

    ws.append(["PHOTO", "CODE", "DESCRIPTION", "PRICE_A INCL"])

    row_idx = 2

    for _, r in df.iterrows():
        temp_img = os.path.join(temp_dir, f"thumb_{row_idx}.jpg")

        # Create 120×120 thumbnail
        img = Image.open(r["TEMP_FILE"])
        img.thumbnail((120, 120))
        img.save(temp_img)

        # Insert thumbnail
        xl_img = XLImage(temp_img)
        xl_img.width = 120
        xl_img.height = 120
        ws.row_dimensions[row_idx].height = 100
        ws.add_image(xl_img, f"A{row_idx}")

        # Add text
        ws[f"B{row_idx}"] = r["CODE"]
        ws[f"C{row_idx}"] = r["DESCRIPTION"]
        ws[f"D{row_idx}"] = r["PRICE_A_INCL"]

        row_idx += 1

    out_path = os.path.join(temp_dir, "catalogue.xlsx")
    wb.save(out_path)
    return out_path

# -----------------------------
# STREAMLIT APP UI
# -----------------------------
def main():
    st.title("Photo Catalogue Builder")
    st.write("Upload Excel (with CODE, DESCRIPTION, PRICE_A_INCL) and product photos.")

    price_file = st.file_uploader("Upload Excel price file", type=["xls", "xlsx"])
    photo_files = st.file_uploader("Upload product photos", type=["jpg", "jpeg", "png"], accept_multiple_files=True)

    if price_file and photo_files:
        if st.button("Generate Catalogue"):
            with tempfile.TemporaryDirectory() as temp_dir:

                # Load price file
                price_df = pd.read_excel(price_file)
                price_df.columns = [c.upper().replace(" ", "_") for c in price_df.columns]

                df = match_photos_to_prices(photo_files, price_df)

                # Save original uploaded bytes for PDF
                file_bytes_list = []
                temp_paths = []

                for i, f in enumerate(photo_files):
                    b = f.read()
                    file_bytes_list.append(b)
                    path = os.path.join(temp_dir, f"orig_{i}.jpg")
                    with open(path, "wb") as out:
                        out.write(b)
                    temp_paths.append(path)

                df["FILE_BYTES"] = file_bytes_list
                df["TEMP_FILE"] = temp_paths

                pdf_bytes = build_pdf(df, temp_dir)
                excel_path = build_excel_with_thumbs(df, temp_dir)

                st.success("Catalogue generated successfully!")

                st.download_button("Download PDF", data=pdf_bytes, file_name="catalogue.pdf")
                st.download_button("Download Excel", data=open(excel_path, "rb").read(), file_name="catalogue.xlsx")

main()
