import streamlit as st
import pandas as pd
import os
import tempfile
from fpdf import FPDF


# ---------------------------------------------------------
# LOAD PRICE FILE (flexible column detection)
# ---------------------------------------------------------
def load_price_file(uploaded_file):
    df = pd.read_excel(uploaded_file, dtype=str)

    df.columns = [c.strip().upper().replace(" ", "").replace("-", "").replace("_", "")
                  for c in df.columns]

    # Attempt to find CODE column
    code_col = None
    for c in df.columns:
        if c.startswith("CODE"):
            code_col = c
            break
    if not code_col:
        raise ValueError("No CODE column found in the price file.")

    # Attempt to find DESCRIPTION column
    desc_col = None
    for c in df.columns:
        if "DESC" in c:
            desc_col = c
            break
    if not desc_col:
        raise ValueError("No DESCRIPTION column found in the price file.")

    # Attempt to find VAT-inclusive price column
    price_col = None
    for c in df.columns:
        if "AINCL" in c or "PRICEAINCL" in c or "PRICEA" in c:
            price_col = c
            break
    if not price_col:
        raise ValueError("Could not find price column (must contain AINCL).")

    df = df[[code_col, desc_col, price_col]].copy()
    df.columns = ["CODE", "DESCRIPTION", "PRICE_A_INCL"]

    df["CODE_STR"] = df["CODE"].astype(str).str.replace(".0", "", regex=False).str.strip()

    df = df.drop_duplicates(subset="CODE_STR", keep="first")

    df["DESCRIPTION"] = df["DESCRIPTION"].fillna("").astype(str)
    df["PRICE_A_INCL"] = df["PRICE_A_INCL"].fillna("").astype(str)

    return df


# ---------------------------------------------------------
# MATCH PHOTOS → CODES
# ---------------------------------------------------------
def extract_base_code(filename):
    base = os.path.splitext(filename)[0]
    digits = "".join([c for c in base if c.isdigit()])
    return digits


def match_photos_to_prices(photo_files, price_df):
    price_dict = price_df.set_index("CODE_STR")[["DESCRIPTION", "PRICE_A_INCL"]].to_dict("index")

    results = []

    for uploaded in photo_files:
        fname = uploaded.name
        code = extract_base_code(fname)
        desc = ""
        price = ""

        if code in price_dict:
            desc = price_dict[code]["DESCRIPTION"]
            price = price_dict[code]["PRICE_A_INCL"]

        results.append({
            "PHOTO_FILE": str(fname),   # <-- FORCE STRING TO FIX ARROW ERROR
            "CODE": code,
            "DESCRIPTION": desc,
            "PRICE_A_INCL": price
        })

    return pd.DataFrame(results)


# ---------------------------------------------------------
# BUILD PDF
# ---------------------------------------------------------
def build_pdf(df, temp_dir):
    pdf = FPDF(unit="mm")
    pdf.set_auto_page_break(auto=True, margin=10)

    pdf.add_page()
    pdf.set_font("Helvetica", "B", 16)
    pdf.cell(0, 10, "Photo Catalogue", 0, 1, "C")

    cell_w = 120
    cell_h = 120

    x_start = 10
    y_start = pdf.get_y() + 5
    x = x_start
    y = y_start
    spacing_y = 10

    for _, row in df.iterrows():
        pdf.set_xy(x, y)

        # SAVE TEMP IMAGE
        temp_path = os.path.join(temp_dir, f"{row['PHOTO_FILE']}.jpg")
        with open(temp_path, "wb") as f:
            f.write(st.session_state["photo_bytes"][row["PHOTO_FILE"]])

        pdf.image(temp_path, x=x, y=y, w=cell_w, h=cell_h)

        y_text = y + cell_h + 3
        pdf.set_xy(x, y_text)
        pdf.set_font("Helvetica", size=10)

        text_block = f"{row['CODE']}\n{row['DESCRIPTION']}\nR {row['PRICE_A_INCL']}"
        for line in text_block.split("\n"):
            pdf.cell(cell_w, 5, line[:60], ln=1)

        y = pdf.get_y() + spacing_y

        if y > 260:
            pdf.add_page()
            y = y_start

    return pdf.output(dest="S").encode("latin1")


# ---------------------------------------------------------
# STREAMLIT APP
# ---------------------------------------------------------
def main():
    st.title("Photo Catalogue Builder — 120×120 Version")

    # Track uploaded photo bytes
    if "photo_bytes" not in st.session_state:
        st.session_state["photo_bytes"] = {}

    price_file = st.file_uploader("Upload price Excel", type=["xls", "xlsx"])
    photo_files = st.file_uploader("Upload product photos", type=["jpg", "jpeg", "png"], accept_multiple_files=True)

    if st.button("Generate Catalogue"):
        if not price_file or not photo_files:
            st.error("Please upload both Excel and photos.")
            return

        df_price = load_price_file(price_file)

        # Save photo bytes
        for f in photo_files:
            st.session_state["photo_bytes"][f.name] = f.read()

        df_matched = match_photos_to_prices(photo_files, df_price)

        st.success("Matched successfully!")

        # SAFE PREVIEW
        preview = df_matched.copy()
        preview["PHOTO_FILE"] = preview["PHOTO_FILE"].astype(str)
        st.dataframe(preview)

        with tempfile.TemporaryDirectory() as temp_dir:
            pdf_bytes = build_pdf(df_matched, temp_dir)
            st.download_button("Download PDF", pdf_bytes, "catalogue.pdf", "application/pdf")

        excel_bytes = df_matched.to_excel(index=False)
        st.download_button("Download Excel", excel_bytes, "catalogue.xlsx", "application/vnd.ms-excel")


if __name__ == "__main__":
    main()
