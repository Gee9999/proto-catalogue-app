import streamlit as st
import pandas as pd
import tempfile
import os
from fpdf import FPDF
from PIL import Image
import io

# --- CLEAN PDF BUILDER (60x60) – unchanged, this was the perfect version ---
def build_pdf(df, temp_dir):
    pdf = FPDF(unit="mm", format="A4")
    pdf.set_auto_page_break(auto=True, margin=10)

    cell_w = 60
    cell_h = 60

    for _, row in df.iterrows():
        photo_file = row["PHOTO_FILE"]
        code = str(row["CODE"])
        desc = str(row["DESCRIPTION"])
        price = str(row["PRICE_A_INCL"])

        img_path = os.path.join(temp_dir, photo_file["name"])
        with open(img_path, "wb") as f:
            f.write(photo_file.getvalue())

        img = Image.open(img_path)
        w, h = img.size
        ratio = w / h
        target_ratio = cell_w / cell_h

        if ratio > target_ratio:
            new_w = cell_w
            new_h = cell_w / ratio
        else:
            new_h = cell_h
            new_w = cell_h * ratio

        pdf.add_page()
        x = (210 - new_w) / 2
        pdf.image(img_path, x=x, y=20, w=new_w, h=new_h)

        # Text under image
        pdf.set_xy(10, 20 + new_h + 10)
        pdf.set_font("Arial", size=10)
        pdf.multi_cell(0, 6, f"CODE: {code}\nDESC: {desc}\nPRICE: {price}")

    # IMPORTANT: return raw bytes, no encode()
    return pdf.output(dest="S")


# --- EXCEL MATCHING ---
def match_photos_to_prices(photo_files, price_df):
    # Normalize price file
    price_df["CODE_STR"] = price_df["CODE"].astype(str)

    price_dict = price_df.set_index("CODE_STR")[["DESCRIPTION", "PRICE_A_INCL"]].to_dict("index")

    rows = []
    for photo in photo_files:
        filename = photo.name.upper()
        numeric_part = "".join([c for c in filename if c.isdigit()])
        code_str = numeric_part[:10]  # first exact match attempt

        if code_str in price_dict:
            matched = price_dict[code_str]
            rows.append({
                "PHOTO_FILE": photo,          # stored for PDF only, NOT for Excel
                "CODE": code_str,
                "DESCRIPTION": matched["DESCRIPTION"],
                "PRICE_A_INCL": matched["PRICE_A_INCL"]
            })
        else:
            rows.append({
                "PHOTO_FILE": photo,
                "CODE": code_str,
                "DESCRIPTION": "",
                "PRICE_A_INCL": ""
            })

    return pd.DataFrame(rows)


# --- MAIN APP ---
def main():
    st.title("Photo Catalogue Builder (PDF + Excel)")

    price_file = st.file_uploader("Upload price Excel", type=["xls", "xlsx"])
    photo_files = st.file_uploader("Upload product photos", type=["jpg", "jpeg", "png"], accept_multiple_files=True)

    if not price_file or not photo_files:
        return

    price_df_raw = pd.read_excel(price_file)

    # Work even if columns are named slightly differently
    rename_map = {}
    for col in price_df_raw.columns:
        c = col.strip().upper().replace(" ", "").replace("-", "")
        if c.startswith("CODE"):
            rename_map[col] = "CODE"
        elif "DESCRIPTION" in c:
            rename_map[col] = "DESCRIPTION"
        elif "PRICE" in c and ("INCL" in c or "AINCL" in c):
            rename_map[col] = "PRICE_A_INCL"

    price_df = price_df_raw.rename(columns=rename_map)

    required = ["CODE", "DESCRIPTION", "PRICE_A_INCL"]
    for r in required:
        if r not in price_df.columns:
            st.error(f"Missing required column: {r}")
            return

    df_matched = match_photos_to_prices(photo_files, price_df)

    # 🔥 FIX: Remove PHOTO_FILE from dataframe before showing/saving Excel
    excel_df = df_matched[["CODE", "DESCRIPTION", "PRICE_A_INCL"]].copy()

    # Show clean table without crashing Streamlit
    st.dataframe(excel_df)

    with tempfile.TemporaryDirectory() as temp_dir:
        # --- Build PDF ---
        pdf_bytes = build_pdf(df_matched, temp_dir)
        st.download_button("Download PDF Catalogue", data=pdf_bytes, file_name="catalogue.pdf", mime="application/pdf")

        # --- Build Excel ---
        excel_buffer = io.BytesIO()
        excel_df.to_excel(excel_buffer, index=False)
        st.download_button("Download Excel File", data=excel_buffer.getvalue(), file_name="catalogue.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


if __name__ == "__main__":
    main()
