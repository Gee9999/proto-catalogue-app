import streamlit as st
import pandas as pd
import tempfile
import os
from fpdf import FPDF
from PIL import Image

# ---------------------------------------------
# Extract numeric code from photo filename
# ---------------------------------------------
def extract_code_from_filename(filename):
    num = ''.join(c for c in filename if c.isdigit())
    return num[:10] if len(num) >= 6 else None


# ---------------------------------------------
# Clean and prepare price Excel file
# ---------------------------------------------
def load_price_file(uploaded_file):

    df = pd.read_excel(uploaded_file, dtype=str)

    df.columns = [c.strip().upper() for c in df.columns]

    # Must contain CODE + DESCRIPTION + PRICE-AINCL.
    required = ["CODE", "DESCRIPTION"]
    for col in required:
        if col not in df.columns:
            raise ValueError(f"Excel must contain column '{col}'")

    # Detect price column
    price_cols = [c for c in df.columns if "AINCL" in c.replace(" ", "")]
    if not price_cols:
        raise ValueError("Could not find price column (PRICE-AINCL.)")
    price_col = price_cols[0]

    df["PRICE_AINCL"] = df[price_col].astype(str)

    # Standardize CODE formatting (remove decimals etc.)
    df["CODE_STR"] = df["CODE"].astype(str).str.replace(".0", "", regex=False).str.strip()

    return df[["CODE_STR", "CODE", "DESCRIPTION", "PRICE_AINCL"]]


# ---------------------------------------------
# Best prefix-based code matching
# ---------------------------------------------
def best_match(photo_code, df):
    exact = df[df["CODE_STR"] == photo_code]
    if not exact.empty:
        return exact.iloc[0]

    # Prefix match (photo code starts with real code)
    df["LEN"] = df["CODE_STR"].str.len()
    subset = df[df["CODE_STR"].apply(lambda x: photo_code.startswith(x))]
    if not subset.empty:
        return subset.sort_values("LEN", ascending=False).iloc[0]

    return None


# ---------------------------------------------
# Match photos to price Excel rows
# ---------------------------------------------
def match_photos(photo_files, df):
    matched_rows = []

    for p in photo_files:
        code = extract_code_from_filename(p.name)
        if not code:
            continue

        row = best_match(code, df)
        if row is None:
            matched_rows.append({
                "PHOTO_FILE": p.name,
                "CODE": code,
                "DESCRIPTION": "NOT FOUND",
                "PRICE_AINCL": "N/A"
            })
        else:
            matched_rows.append({
                "PHOTO_FILE": p.name,
                "CODE": row["CODE_STR"],
                "DESCRIPTION": row["DESCRIPTION"],
                "PRICE_AINCL": row["PRICE_AINCL"]
            })

    return pd.DataFrame(matched_rows)


# ---------------------------------------------
# Build PDF (3×3 layout, perfect spacing)
# ---------------------------------------------
def build_pdf(df, temp_dir, thumb_size=120):

    class PDF(FPDF):
        pass

    pdf = PDF()
    pdf.set_auto_page_break(auto=True, margin=10)
    pdf.add_page()

    cols = 3
    cell_w = 65
    cell_h = 65

    x_start = 10
    y_start = 20
    x = x_start
    y = y_start

    count = 0

    for _, row in df.iterrows():
        img_path = os.path.join(temp_dir, "_tmp_" + row["PHOTO_FILE"])

        # Process thumbnail
        try:
            img = Image.open(os.path.join(temp_dir, row["PHOTO_FILE"]))
            img.thumbnail((thumb_size, thumb_size))
            img.save(img_path)
        except:
            continue

        # Draw image
        pdf.image(img_path, x=x, y=y, w=cell_w)

        # Move text UNDER the image
        text_y = y + thumb_size + 2
        pdf.set_xy(x, text_y)
        pdf.set_font("Arial", size=8)

        pdf.multi_cell(cell_w, 4, f"{row['CODE']}\n{row['DESCRIPTION']}\nR {row['PRICE_AINCL']}", 0, "L")

        count += 1
        x += cell_w + 5

        # New row
        if count % cols == 0:
            x = x_start
            y += thumb_size + 18

        # New page
        if y > 240:
            pdf.add_page()
            x = x_start
            y = y_start

    return pdf.output(dest="S").encode("latin1")


# ---------------------------------------------
# Build Excel with thumbnails (120×120)
# ---------------------------------------------
def build_excel(df, temp_dir, thumb_size=120):
    from openpyxl import Workbook
    from openpyxl.drawing.image import Image as XLImage

    wb = Workbook()
    ws = wb.active

    ws.append(["PHOTO", "CODE", "DESCRIPTION", "PRICE_AINCL"])

    for idx, row in df.iterrows():
        ws.append(["", row["CODE"], row["DESCRIPTION"], row["PRICE_AINCL"]])

        img_path = os.path.join(temp_dir, row["PHOTO_FILE"])
        thumb_path = img_path + "_xl.jpg"

        try:
            img = Image.open(img_path)
            img.thumbnail((thumb_size, thumb_size))
            img.save(thumb_path)

            xl_img = XLImage(thumb_path)
            xl_img.width = thumb_size
            xl_img.height = thumb_size

            ws.add_image(xl_img, f"A{idx+2}")

            ws.row_dimensions[idx+2].height = thumb_size * 0.75

        except:
            pass

    out = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    wb.save(out.name)
    out.seek(0)
    return out.read()


# ---------------------------------------------
# Streamlit App
# ---------------------------------------------
def main():

    st.title("📸 Photo Catalogue Builder")

    st.write("Upload price Excel + photos → get PDF + Excel")

    price_file = st.file_uploader("Upload price Excel", type=["xlsx", "xls"])
    photos = st.file_uploader("Upload product photos", type=["jpg", "jpeg", "png"], accept_multiple_files=True)

    if not price_file or not photos:
        return

    # Load price sheet
    df_price = load_price_file(price_file)

    # Temp directory
    with tempfile.TemporaryDirectory() as temp_dir:

        # Save photos
        for p in photos:
            with open(os.path.join(temp_dir, p.name), "wb") as f:
                f.write(p.getbuffer())

        # Match photos to price Excel
        df_match = match_photos(photos, df_price)

        # PDF & Excel
        pdf_bytes = build_pdf(df_match, temp_dir, thumb_size=120)
        excel_bytes = build_excel(df_match, temp_dir, thumb_size=120)

        st.success("Done! Download your files:")

        st.download_button("📄 Download PDF", data=pdf_bytes, file_name="catalogue.pdf")
        st.download_button("📊 Download Excel", data=excel_bytes, file_name="catalogue.xlsx")


if __name__ == "__main__":
    main()
