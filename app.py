import streamlit as st
import pdfplumber
import pandas as pd
import re
import io
import os
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader

# ---- 1. PRICE EXTRACTION FROM PDF ----
def extract_prices_from_pdf(pdf_file):
    """
    Reads a Proto Trading price PDF and extracts:
    - CODE (string from PDF)
    - BASE_CODE (digits-only part of CODE, for matching with photos)
    - DESCRIPTION
    - PRICE_INCL (Price-A Incl, 5th decimal number on the line)
    """
    items = []
    code_pattern = re.compile(r"^(\d+[A-Za-z]?)\b") # e.g. 8610100002, 8610100041G

    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if not text:
                continue

            for line in text.split("\n"):
                line = line.strip()
                if not line:
                    continue

                if not code_pattern.match(line):
                    continue

                parts = line.split()
                code = parts[0]

                numbers = [p for p in parts if re.match(r"^\d+\.\d+$", p)]
                if len(numbers) < 5:
                    continue

                price_incl = float(numbers[4])

                desc_tokens = []
                for p in parts[1:]:
                    if re.match(r"^\d+\.\d+$", p):
                        break
                    desc_tokens.append(p)
                description = " ".join(desc_tokens)

                # base code = digits-only from start of CODE
                m = re.match(r"^(\d+)", code)
                base_code = m.group(1) if m else code

                items.append(
                    {
                        "CODE": code,
                        "BASE_CODE": base_code,
                        "DESCRIPTION": description,
                        "PRICE_INCL": price_incl,
                    }
                )
    df = pd.DataFrame(items, columns=["CODE", "BASE_CODE", "DESCRIPTION", "PRICE_INCL"])
    return df

# ---- 2. BUILD MATCHED TABLE (PHOTOS + PRICES) ----
def build_photo_price_table(files, price_df):
    """
    files: list of UploadedFile objects from st.file_uploader
    price_df: DataFrame from extract_prices_from_pdf
    Returns a DataFrame with one row per photo:
    - FILENAME
    - BASE_CODE (from filename)
    - CODE (from PDF, if matched)
    - DESCRIPTION
    - PRICE_INCL
    """
    rows = []
    for f in files:
        filename = f.name
        stem, _ = os.path.splitext(filename)

        # base code from filename = leading digits, remove spaces/dashes first
        cleaned_stem = stem.replace(" ", "").replace("-", "")
        m = re.match(r"^(\d+)", cleaned_stem)
        base_code = m.group(1) if m else None

        match_row = None
        if base_code is not None and not price_df.empty:
            matches = price_df[price_df["BASE_CODE"] == base_code]
            if not matches.empty:
                match_row = matches.iloc[0]

        rows.append(
            {
                "FILENAME": filename,
                "BASE_CODE": base_code,
                "CODE": match_row["CODE"] if match_row is not None else None,
                "DESCRIPTION": match_row["DESCRIPTION"] if match_row is not None else None,
                "PRICE_INCL": match_row["PRICE_INCL"] if match_row is not None else None,
            }
        )
    return pd.DataFrame(rows)

# ---- 3. CONVERT DF TO EXCEL BYTES ----
def df_to_excel_bytes(df: pd.DataFrame, sheet_name: str = "Sheet1") -> bytes:
    """ Convert a DataFrame to an in-memory Excel file for download in Streamlit. """
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
    buffer.seek(0)
    return buffer

# ---- 4. GENERATE PDF GRID (4 ROWS x 2 COLS) ----
def generate_grid_pdf(files, matched_df) -> bytes:
    """
    Creates a PDF with 4 rows x 2 columns per page.
    Each cell shows the image, code, description, and price.
    """
    buffer = io.BytesIO()
    c = canvas.Canvas(buffer, pagesize=A4)
    page_width, page_height = A4
    margin_x = 30
    margin_y = 40
    cols = 2
    rows = 4
    cell_width = (page_width - 2 * margin_x) / cols
    cell_height = (page_height - 2 * margin_y) / rows

    # Mapping from filename to uploaded file
    file_dict = {f.name: f for f in files}

    def draw_item(idx, row):
        col_idx = idx % cols
        row_idx = (idx // cols) % rows

        x0 = margin_x + col_idx * cell_width
        y0 = page_height - margin_y - (row_idx + 1) * cell_height

        filename = row["FILENAME"]
        code = row.get("CODE") or row.get("BASE_CODE") or ""
        desc = row.get("DESCRIPTION") or ""
        price = row.get("PRICE_INCL")

        img_file = file_dict.get(filename)
        img_max_width = cell_width - 10
        img_max_height = cell_height * 0.55

        if img_file is not None:
            img_data = img_file.getvalue()
            try:
                img = ImageReader(io.BytesIO(img_data))
                iw, ih = img.getSize()
                scale = min(img_max_width / iw, img_max_height / ih)
                img_w = iw * scale
                img_h = ih * scale
                img_x = x0 + (cell_width - img_w) / 2
                img_y = y0 + cell_height - img_h - 5
                c.drawImage(img, img_x, img_y, img_w, img_h, preserveAspectRatio=True, mask="auto")
            except Exception:
                pass # Handle cases where image might be corrupted or unreadable

        text_x = x0 + 5
        text_y = y0 + cell_height * 0.4

        c.setFont("Helvetica-Bold", 9)
        c.drawString(text_x, text_y, f"Code: {code}")
        c.setFont("Helvetica", 8)
        c.drawString(text_x, text_y - 12, f"{desc[:70]}") # Truncate long descriptions

        if price is not None and not pd.isna(price):
            c.setFont("Helvetica-Bold", 9)
            c.drawString(text_x, text_y - 26, f"Price (incl): R {price:0.2f}")
        else:
            c.setFont("Helvetica-Oblique", 8)
            c.drawString(text_x, text_y - 26, "Price not found")

    total_items = len(matched_df)
    items_per_page = cols * rows

    for page_start in range(0, total_items, items_per_page):
        page_slice = matched_df.iloc[page_start : page_start + items_per_page]
        for i, (_, row) in enumerate(page_slice.iterrows()):
            draw_item(i, row)
        c.showPage() # Start a new page after every 8 items

    c.save()
    buffer.seek(0)
    return buffer

# ---- 5. STREAMLIT APP UI ----
def main():
    st.set_page_config(page_title="Photo + Price PDF Exporter", layout="wide")
    st.title("📸 Photo to Price Sheet & PDF Exporter")
    st.write(
        "Upload your **price PDF** and **product photos**.\n\n"
        "I'll match photos to product codes (based on leading digits in the filename), "
        "pull **Price-A Incl** from the PDF, and generate:\n"
        "- an Excel sheet, and\n"
        "- a printable PDF with **4 rows × 2 columns** per page."
    )

    st.subheader("1️⃣ Upload price PDF")
    price_pdf = st.file_uploader("Price PDF", type=["pdf"])

    st.subheader("2️⃣ Upload product photos")
    photo_files = st.file_uploader(
        "Product photos (jpg / jpeg / png)",
        type=["jpg", "jpeg", "png"],
        accept_multiple_files=True,
    )

    if price_pdf is not None and photo_files:
        if st.button("🔍 Process & Generate Outputs"):
            with st.spinner("Extracting prices from PDF…"):
                price_df = extract_prices_from_pdf(price_pdf)
                if price_df.empty:
                    st.error("No items found in the price PDF. Check that it's the correct report.")
                    return
                st.success(f"Loaded {len(price_df)} items from the price PDF.")

            with st.spinner("Matching photos with codes & prices…"):
                matched_df = build_photo_price_table(photo_files, price_df)

            st.subheader("3️⃣ Preview matched data")
            st.dataframe(matched_df, use_container_width=True)

            # Excel download
            excel_bytes = df_to_excel_bytes(matched_df, sheet_name="PhotoPrices")
            st.download_button(
                label="⬇️ Download Excel (photo_price_output.xlsx)",
                data=excel_bytes,
                file_name="photo_price_output.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

            # PDF download
            with st.spinner("Generating PDF (4 rows × 2 columns)…"):
                pdf_bytes = generate_grid_pdf(photo_files, matched_df)
                st.download_button(
                    label="⬇️ Download PDF (photo_catalogue.pdf)",
                    data=pdf_bytes,
                    file_name="photo_catalogue.pdf",
                    mime="application/pdf",
                )
    else:
        st.info("Upload both a price PDF and at least one product photo to continue.")

if __name__ == "__main__":
    main()