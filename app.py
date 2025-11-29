import streamlit as st
import pandas as pd
import tempfile
from PIL import Image
import io
import os

st.set_page_config(page_title="Photo Catalogue Builder", layout="wide")

# ---------------------------
# 1) Extract numeric code from photo filename
# ---------------------------
def extract_code_from_filename(filename):
    digits = "".join([c for c in filename if c.isdigit()])
    return digits[:11] if len(digits) >= 11 else digits

# ---------------------------
# 2) Load Excel price file and auto-detect columns
# ---------------------------
def load_price_file(file):
    df = pd.read_excel(file, dtype=str)

    df.columns = df.columns.str.upper().str.replace(" ", "").str.replace("-", "")

    # Auto-detect
    code_col = None
    desc_col = None
    price_col = None

    for col in df.columns:
        if "CODE" in col:
            code_col = col
        if "DESC" in col:
            desc_col = col
        if ("PRICE" in col or "AINCL" in col or ("INCL" in col and "EXCL" not in col)):
            price_col = col

    if not code_col or not desc_col or not price_col:
        st.error(f"""
        ❌ Could not auto-detect columns.
        Found columns: {df.columns.tolist()}
        Need something that looks like:
        - CODE
        - DESCRIPTION
        - PRICE-A INCL
        """)
        st.stop()

    df = df.rename(columns={
        code_col: "CODE",
        desc_col: "DESCRIPTION",
        price_col: "PRICE_A_INCL"
    })

    df["CODE_STR"] = df["CODE"].astype(str).str.replace(r"\D", "", regex=True)
    df = df.dropna(subset=["CODE_STR"])

    return df[["CODE_STR", "DESCRIPTION", "PRICE_A_INCL", "CODE"]]

# ---------------------------
# 3) Match photos to codes
# ---------------------------
def match_photos_to_prices(photo_files, price_df):
    price_df = price_df.drop_duplicates(subset=["CODE_STR"])
    price_dict = price_df.set_index("CODE_STR")[["DESCRIPTION", "PRICE_A_INCL", "CODE"]].to_dict("index")

    rows = []

    for pf in photo_files:
        filename = pf.name
        code_from_photo = extract_code_from_filename(filename)

        desc = ""
        price = ""
        code_exact = ""

        if code_from_photo in price_dict:
            desc = price_dict[code_from_photo]["DESCRIPTION"]
            price = price_dict[code_from_photo]["PRICE_A_INCL"]
            code_exact = price_dict[code_from_photo]["CODE"]

        rows.append({
            "FILENAME": filename,
            "CODE": code_exact,
            "DESCRIPTION": desc,
            "PRICE_A_INCL": price,
            "PHOTO_FILE": pf
        })

    return pd.DataFrame(rows)

# ---------------------------
# 4) Build Excel with thumbnails
# ---------------------------
def build_excel_with_thumbnails(df, temp_dir):
    from openpyxl import Workbook
    from openpyxl.drawing.image import Image as XLImage

    wb = Workbook()
    ws = wb.active
    ws.title = "Catalogue"

    ws.append(["PHOTO", "CODE", "DESCRIPTION", "PRICE_A_INCL", "FILENAME"])

    thumb_size = (120, 120)

    for idx, row in df.iterrows():
        photo = row["PHOTO_FILE"]

        img = Image.open(photo)
        img.thumbnail(thumb_size)

        temp_path = os.path.join(temp_dir, f"thumb_{idx}.png")
        img.save(temp_path)

        ws.append(["", row["CODE"], row["DESCRIPTION"], row["PRICE_A_INCL"], row["FILENAME"]])

        xl_img = XLImage(temp_path)
        xl_img.width, xl_img.height = thumb_size

        ws.row_dimensions[idx + 2].height = 100
        ws.column_dimensions["A"].width = 20

        ws.add_image(xl_img, f"A{idx + 2}")

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out


# ---------------------------
# MAIN APP
# ---------------------------
def main():
    st.title("📸 Photo Catalogue Builder — Excel Version")

    st.subheader("1️⃣ Upload price Excel file")
    price_file = st.file_uploader("Upload Excel", type=["xls", "xlsx"])

    st.subheader("2️⃣ Upload product photos")
    photos = st.file_uploader("Upload photos", type=["jpg", "jpeg", "png"], accept_multiple_files=True)

    if price_file and photos:
        st.success("Files uploaded successfully ✔")

        df_price = load_price_file(price_file)

        st.info("Matching photos to Excel...")
        df_matched = match_photos_to_prices(photos, df_price)

        st.subheader("3️⃣ Generate Excel")

        if st.button("Generate Excel Catalogue"):
            with tempfile.TemporaryDirectory() as temp_dir:
                excel_bytes = build_excel_with_thumbnails(df_matched, temp_dir)

            st.success("Excel created successfully!")

            st.download_button(
                "Download Excel Catalogue",
                excel_bytes,
                file_name="photo_catalogue.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

            st.dataframe(df_matched)


if __name__ == "__main__":
    main()
