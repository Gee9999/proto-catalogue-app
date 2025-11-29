import streamlit as st
import pandas as pd
import tempfile
import os
import re
from fpdf import FPDF
from difflib import SequenceMatcher

# ------------------------------------------
#         PDF BUILDER (final + stable)
# ------------------------------------------
def build_pdf(matched_df, temp_dir):
    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=10)

    pdf.add_page()
    pdf.set_font("Helvetica", "B", 14)
    pdf.cell(0, 10, "Photo Catalogue", new_x="LMARGIN", new_y="NEXT", align="C")

    cell_w = 90
    cell_h = 60

    # Build grid: 2 photos per row
    for idx, row in matched_df.iterrows():
        if idx % 2 == 0:
            pdf.ln(5)

        x = pdf.get_x()
        y = pdf.get_y()

        photo_path = os.path.join(temp_dir, row["PHOTO_FILE"])
        if os.path.exists(photo_path):
            pdf.image(photo_path, x=x, y=y, w=cell_w, h=cell_h)

        pdf.set_xy(x, y + cell_h + 2)
        pdf.set_font("Helvetica", size=9)

        pdf.multi_cell(cell_w, 4, f"CODE: {row['CODE']}")
        pdf.multi_cell(cell_w, 4, f"DESC: {row['DESCRIPTION']}")
        pdf.multi_cell(cell_w, 4, f"PRICE: {row['PRICE_A_INCL']}")

        if idx % 2 == 0:
            pdf.set_xy(x + cell_w + 10, y)

    return pdf.output(dest="S")

# ------------------------------------------
#          CLEAN COLUMN NAME FINDER
# ------------------------------------------
def find_column(df, targets):
    cols = {c.lower().replace(" ", "").replace("-", ""): c for c in df.columns}
    for t in targets:
        t_clean = t.lower().replace(" ", "").replace("-", "")
        for key in cols:
            if t_clean in key:
                return cols[key]
    return None

# ------------------------------------------
#          MATCHING FUNCTION
# ------------------------------------------
def extract_code_from_filename(fname):
    nums = re.findall(r"\d+", fname)
    return max(nums, key=len) if nums else ""

def match_code_to_price(code, price_df):
    if code in price_df["CODE_STR"].values:
        return code

    best = None
    best_score = 0.0
    for c in price_df["CODE_STR"].values:
        score = SequenceMatcher(None, code, c).ratio()
        if score > best_score:
            best = c
            best_score = score

    return best

def match_photos_to_prices(photo_filenames, price_df):
    rows = []

    for fname in photo_filenames:
        extracted_code = extract_code_from_filename(fname)
        best_code = match_code_to_price(extracted_code, price_df)

        match = price_df[price_df["CODE_STR"] == best_code].iloc[0]

        rows.append({
            "PHOTO_FILE": fname,
            "CODE": match["CODE"],
            "DESCRIPTION": match["DESCRIPTION"],
            "PRICE_A_INCL": match["PRICE_A_INCL"]
        })

    return pd.DataFrame(rows)

# ------------------------------------------
#               MAIN STREAMLIT APP
# ------------------------------------------
def main():
    st.title("Photo Catalogue Builder")

    st.write("Upload **any price Excel** containing CODE + DESCRIPTION + PRICE-AINCL.")
    price_file = st.file_uploader("Upload price Excel", type=["xls", "xlsx"])

    st.write("Upload product **photos** (JPG/PNG).")
    photo_files = st.file_uploader("Upload product photos", type=["jpg", "jpeg", "png"], accept_multiple_files=True)

    if not price_file or not photo_files:
        return

    if st.button("Generate Catalogue"):
        with st.spinner("Processing..."):
            try:
                # ----- Create temp dir -----
                with tempfile.TemporaryDirectory() as tmpdir:
                    df = pd.read_excel(price_file)

                    # --- Identify columns ---
                    code_col = find_column(df, ["code"])
                    desc_col = find_column(df, ["description", "desc"])
                    price_col = find_column(df, ["priceaincl", "priceincl", "price"])

                    if not code_col or not desc_col or not price_col:
                        st.error("Excel must contain CODE, DESCRIPTION and PRICE-AINCL columns.")
                        return

                    df["CODE_STR"] = df[code_col].astype(str).str.extract(r"(\d+)")
                    df["PRICE_A_INCL"] = df[price_col]

                    # ----- Save photos -----
                    saved_filenames = []
                    for pf in photo_files:
                        fname = pf.name
                        out_path = os.path.join(tmpdir, fname)
                        with open(out_path, "wb") as f:
                            f.write(pf.read())
                        saved_filenames.append(fname)

                    # ----- Match -----
                    matched_df = match_photos_to_prices(saved_filenames, df)

                    # ----- Build PDF -----
                    pdf_bytes = build_pdf(matched_df, tmpdir)

                    # ----- Build Excel -----
                    excel_path = os.path.join(tmpdir, "catalogue.xlsx")
                    matched_df.to_excel(excel_path, index=False)
                    with open(excel_path, "rb") as f:
                        excel_bytes = f.read()

                    # ----- Outputs -----
                    st.success("Catalogue built!")

                    st.download_button("Download PDF", pdf_bytes, file_name="catalogue.pdf")
                    st.download_button("Download Excel", excel_bytes, file_name="catalogue.xlsx")

            except Exception as e:
                st.error(f"Something went wrong: {e}")

# ------------------------------------------
# ENTRY POINT
# ------------------------------------------
if __name__ == "__main__":
    main()
