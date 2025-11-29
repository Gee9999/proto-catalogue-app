import streamlit as st
import pandas as pd
import tempfile
import os
from PIL import Image
from fpdf import FPDF

# ============================================================
# 1. EXTRACT NUMERIC CODE FROM PHOTO FILENAME
# ============================================================
def extract_code_from_filename(filename):
    base = os.path.splitext(filename)[0]
    numbers = "".join([c for c in base if c.isdigit()])
    if len(numbers) < 4:
        return None
    return numbers


# ============================================================
# 2. CLEAN EXCEL COLUMNS (AUTO-DETECT CODE, DESCRIPTION, PRICE)
# ============================================================
def load_price_excel(uploaded_file):
    df = pd.read_excel(uploaded_file, dtype=str)

    df.columns = df.columns.str.strip().str.upper()

    # Required columns
    code_col = None
    desc_col = None
    price_col = None

    for col in df.columns:
        if "CODE" in col:
            code_col = col
        if "DESCRIPTION" in col or "DESC" in col:
            desc_col = col
        if "PRICE" in col and "INCL" in col:
            price_col = col

    if not code_col or not desc_col or not price_col:
        st.error("Excel must contain CODE, DESCRIPTION and PRICE-AINCL columns.")
        return None

    df = df[[code_col, desc_col, price_col]]
    df.columns = ["CODE", "DESCRIPTION", "PRICE_A_INCL"]

    # Cleanup
    df["CODE"] = df["CODE"].astype(str).str.replace(r"\D", "", regex=True)
    df["PRICE_A_INCL"] = (
        df["PRICE_A_INCL"]
        .astype(str)
        .str.replace(",", ".")
        .str.extract(r"([0-9]+\.[0-9]+|[0-9]+)")
    )

    df["PRICE_A_INCL"] = pd.to_numeric(df["PRICE_A_INCL"], errors="coerce")

    return df


# ============================================================
# 3. MATCH PHOTOS TO EXCEL ROWS PERFECTLY
# ============================================================
def match_photos_to_prices(photo_files, price_df):
    price_df["CODE_STR"] = price_df["CODE"].astype(str)

    lookup = price_df.set_index("CODE_STR")[["CODE", "DESCRIPTION", "PRICE_A_INCL"]]

    results = []

    for file in photo_files:
        filename = file.name
        extracted = extract_code_from_filename(filename)
        info = lookup.loc[extracted] if extracted in lookup.index else None

        results.append({
            "PHOTO_FILE": filename,
            "EXTRACTED_CODE": extracted,
            "CODE": info["CODE"] if info is not None else "",
            "DESCRIPTION": info["DESCRIPTION"] if info is not None else "",
            "PRICE_A_INCL": info["PRICE_A_INCL"] if info is not None else ""
        })

    return pd.DataFrame(results)


# ============================================================
# 4. PDF GENERATION — 3×3 GRID WITH FIXED IMAGE HEIGHT
# ============================================================
def build_pdf(df, temp_dir):
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

        img_path = item.get("TEMP_PATH")
        if img_path and os.path.exists(img_path):
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

        # --- TEXT ALWAYS BELOW IMAGE ---
        pdf.set_xy(x + 2, y + fixed_image_height + 2)
        pdf.set_font("Arial", size=8)

        code_line = f"Code: {item['CODE']}" if item["CODE"] else f"Code: {item['PHOTO_FILE']}"
        desc_line = str(item["DESCRIPTION"])[:60]
        price_line = (
            f"Price: {item['PRICE_A_INCL']:,.2f}"
            if isinstance(item["PRICE_A_INCL"], (int, float))
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

    return pdf.output(dest="S").encode("latin1")


# ============================================================
# 5. BUILD EXCEL WITH THUMBNAILS
# ============================================================
def build_excel(df, temp_dir):
    out = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    writer = pd.ExcelWriter(out.name, engine="openpyxl")

    df[["PHOTO_FILE", "CODE", "DESCRIPTION", "PRICE_A_INCL"]].to_excel(
        writer, index=False, sheet_name="Catalogue"
    )
    writer.save()

    with open(out.name, "rb") as f:
        return f.read()


# ============================================================
# 6. MAIN STREAMLIT UI
# ============================================================
def main():
    st.title("Photo Catalogue Builder")

    price_file = st.file_uploader("Upload price Excel", type=["xlsx", "xls"])
    photo_files = st.file_uploader(
        "Upload product photos", type=["jpg", "jpeg", "png"], accept_multiple_files=True
    )

    if st.button("Generate catalogue"):
        if not price_file or not photo_files:
            st.error("Upload Excel + Photos first.")
            return

        price_df = load_price_excel(price_file)
        if price_df is None:
            return

        matched_df = match_photos_to_prices(photo_files, price_df)

        temp_dir = tempfile.mkdtemp()
        for i, file in enumerate(photo_files):
            img = Image.open(file)
            resized = img.resize((300, 300))
            temp_path = os.path.join(temp_dir, f"img_{i}.jpg")
            resized.save(temp_path)
            matched_df.loc[i, "TEMP_PATH"] = temp_path

        pdf_bytes = build_pdf(matched_df, temp_dir)
        excel_bytes = build_excel(matched_df, temp_dir)

        st.success("Catalogue created!")

        st.download_button(
            "Download Catalogue PDF",
            pdf_bytes,
            file_name="catalogue.pdf",
            mime="application/pdf"
        )

        st.download_button(
            "Download Excel",
            excel_bytes,
            file_name="catalogue.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )


if __name__ == "__main__":
    main()
