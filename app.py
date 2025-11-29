import os
import io
import re
import math
import tempfile

import pandas as pd
from fpdf import FPDF
from PIL import Image
import streamlit as st


def normalize_header(name: str) -> str:
    """Normalize Excel header names so we can detect CODE / DESCRIPTION / PRICE columns."""
    if not isinstance(name, str):
        name = str(name)
    s = name.strip().lower()
    # Excel sometimes encodes newlines as _x000a_
    s = s.replace("_x000a_", " ").replace("x000a", "")
    # keep only letters and digits
    return re.sub(r"[^a-z0-9]", "", s)


def load_price_excel(uploaded_file) -> pd.DataFrame:
    """Read the price Excel and return a DataFrame with columns CODE, DESCRIPTION, PRICE_A_INCL."""
    df_raw = pd.read_excel(uploaded_file)

    norms = {col: normalize_header(col) for col in df_raw.columns}

    def find_col(candidates):
        for col, norm in norms.items():
            if norm in candidates:
                return col
        return None

    code_col = find_col({"code", "itemcode", "stockcode", "productcode", "barcode"})
    desc_col = find_col({"description", "itemdescription", "productdescription", "desc"})
    price_col = find_col(
        {
            "priceaincl",
            "priceincl",
            "sellingpriceincl",
            "sellincl",
            "priceainecl",
            "pricewithvat",
        }
    )

    if not code_col or not desc_col or not price_col:
        raise ValueError(
            "Could not automatically find CODE, DESCRIPTION and PRICE-A INCL columns "
            "in the Excel file. Please make sure there is one column for each."
        )

    df = df_raw[[code_col, desc_col, price_col]].copy()
    df.columns = ["CODE", "DESCRIPTION", "PRICE_A_INCL"]

    # Clean up text columns
    df["CODE"] = df["CODE"].astype(str).str.strip()
    df["DESCRIPTION"] = df["DESCRIPTION"].astype(str).str.strip()

    # Clean up price: remove spaces / commas, coerce to number
    df["PRICE_A_INCL"] = (
        df["PRICE_A_INCL"]
        .astype(str)
        .str.replace(",", "", regex=False)
        .str.replace(" ", "", regex=False)
    )
    df["PRICE_A_INCL"] = pd.to_numeric(df["PRICE_A_INCL"], errors="coerce")

    return df


def normalize_code_for_match(code) -> str:
    """Turn the Excel CODE into a pure digit string so it matches photo filenames."""
    if pd.isna(code):
        return ""
    s = str(code).strip()
    # drop trailing .0 from numbers
    if s.endswith(".0"):
        s = s[:-2]
    # keep only digits
    digits = re.sub(r"\D", "", s)
    return digits or s


def extract_base_code_from_filename(filename: str) -> str | None:
    """Extract the longest digit group from the photo filename."""
    base = os.path.basename(filename)
    nums = re.findall(r"\d+", base)
    if not nums:
        return None
    return max(nums, key=len)


def match_photos_to_prices(photo_files, price_df: pd.DataFrame) -> pd.DataFrame:
    """Return df with PHOTO_FILE, CODE, DESCRIPTION, PRICE_A_INCL."""
    df = price_df.copy()
    df["CODE_STR"] = df["CODE"].apply(normalize_code_for_match)

    # If there are duplicate codes, keep the first – avoids index-uniqueness errors.
    df_unique = df.drop_duplicates(subset="CODE_STR", keep="first")

    rows = []
    for pf in photo_files:
        fname = pf.name if hasattr(pf, "name") else str(pf)
        photo_code = extract_base_code_from_filename(fname)
        if not photo_code:
            continue

        # Exact match on digit string
        subset = df_unique[df_unique["CODE_STR"] == photo_code]

        # Fallback: where the Excel code contains the photo code (for partial matches)
        if subset.empty:
            subset = df_unique[df_unique["CODE_STR"].str.contains(photo_code, na=False)]

        if subset.empty:
            continue

        row = subset.iloc[0]
        rows.append(
            {
                "PHOTO_FILE": fname,
                "CODE": row["CODE"],
                "DESCRIPTION": row["DESCRIPTION"],
                "PRICE_A_INCL": row["PRICE_A_INCL"],
            }
        )

    return pd.DataFrame(rows)


def build_pdf(matched_df: pd.DataFrame, temp_dir: str) -> bytes:
    """Create the photo catalogue PDF. Text is always UNDER the photo."""
    pdf = FPDF("P", "mm", "A4")
    pdf.set_auto_page_break(auto=False, margin=10)
    pdf.add_page()

    # Title
    pdf.set_font("Arial", "B", 14)
    pdf.cell(0, 10, "Photo Catalogue", ln=1, align="C")
    pdf.ln(2)

    # Layout: 3 columns x 4 rows per page
    cols = 3
    rows = 4
    margin_x = 10
    margin_top = 20
    gutter_x = 5
    gutter_y = 8

    usable_width = pdf.w - 2 * margin_x
    cell_w = (usable_width - (cols - 1) * gutter_x) / cols
    cell_h = (pdf.h - margin_top - 15 - (rows - 1) * gutter_y) / rows

    # Reserve space for text
    max_img_h = cell_h - 18

    pdf.set_font("Arial", size=8)

    for idx, row in matched_df.iterrows():
        grid_index = idx % (cols * rows)

        if grid_index == 0 and idx != 0:
            pdf.add_page()
            pdf.set_font("Arial", "B", 14)
            pdf.cell(0, 10, "Photo Catalogue", ln=1, align="C")
            pdf.ln(2)
            pdf.set_font("Arial", size=8)

        row_pos = grid_index // cols
        col_pos = grid_index % cols

        x0 = margin_x + col_pos * (cell_w + gutter_x)
        y0 = margin_top + row_pos * (cell_h + gutter_y)

        img_path = os.path.join(temp_dir, row["PHOTO_FILE"])

        # Draw image, scaled to fit in the top part of the cell
        try:
            with Image.open(img_path) as im:
                w, h = im.size
                # start with width-based scaling
                img_w = cell_w * 0.9
                img_h = img_w * h / w
                if img_h > max_img_h:
                    img_h = max_img_h
                    img_w = img_h * w / h
        except Exception:
            # If image can't be opened, just skip drawing it but still print text
            img_w = 0
            img_h = 0

        if img_w > 0 and img_h > 0:
            img_x = x0 + (cell_w - img_w) / 2
            img_y = y0
            pdf.image(img_path, x=img_x, y=img_y, w=img_w, h=img_h)

        # Text UNDER the photo
        text_y = y0 + img_h + 2
        pdf.set_xy(x0, text_y)

        price_val = row["PRICE_A_INCL"]
        if pd.isna(price_val):
            price_str = ""
        else:
            price_str = f"R{price_val:,.2f}"

        code_line = f"CODE: {row['CODE']}"
        desc_line = f"{row['DESCRIPTION']}"
        price_line = f"PRICE-A INCL: {price_str}"

        # Make sure we never write text above the image
        pdf.set_xy(x0, text_y)
        pdf.multi_cell(cell_w, 4, code_line, align="L")
        pdf.set_xy(x0, text_y + 4)
        pdf.multi_cell(cell_w, 4, desc_line, align="L")
        pdf.set_xy(x0, text_y + 8)
        pdf.multi_cell(cell_w, 4, price_line, align="L")

    raw = pdf.output(dest="S")
    if isinstance(raw, bytearray):
        raw = bytes(raw)
    return raw


def build_excel(matched_df: pd.DataFrame) -> bytes:
    """Return Excel file as bytes with PHOTO_FILE, CODE, DESCRIPTION, PRICE-A INCL."""
    output = io.BytesIO()
    out_df = matched_df[["PHOTO_FILE", "CODE", "DESCRIPTION", "PRICE_A_INCL"]].copy()
    out_df.rename(columns={"PRICE_A_INCL": "PRICE-A INCL."}, inplace=True)
    out_df.to_excel(output, index=False)
    output.seek(0)
    return output.read()


def main():
    st.set_page_config(page_title="Photo Catalogue Builder", layout="wide")

    st.title("Photo Catalogue Builder")

    st.markdown(
        """
Upload your **price Excel** file (any layout, as long as it has a CODE, DESCRIPTION and VAT-inclusive price column).  
Upload the **product photos** (filenames must contain the product code, e.g. `8613900001-25pcs.jpg`).  

Click **Generate catalogue** to get:

- A **PDF photo catalogue**
- A **matching Excel** file with columns `PHOTO_FILE, CODE, DESCRIPTION, PRICE-A INCL.`
"""
    )

    price_file = st.file_uploader(
        "Upload price Excel (e.g. PRODUCT DETAILS - BY CODE.xlsx)",
        type=["xls", "xlsx"],
    )

    photo_files = st.file_uploader(
        "Upload product photos", type=["jpg", "jpeg", "png"], accept_multiple_files=True
    )

    if price_file is not None:
        st.caption(f"Price file: {price_file.name}")
    if photo_files:
        st.caption(f"{len(photo_files)} photo(s) uploaded")

    if st.button("Generate catalogue"):
        if not price_file or not photo_files:
            st.error("Please upload both a price Excel file and at least one product photo.")
            return

        try:
            price_df = load_price_excel(price_file)
        except Exception as e:
            st.error(f"Could not read price Excel: {e}")
            return

        with tempfile.TemporaryDirectory() as tmpdir:
            # Save photos temporarily so FPDF can read them by path
            for pf in photo_files:
                img_bytes = pf.read()
                out_path = os.path.join(tmpdir, pf.name)
                with open(out_path, "wb") as f:
                    f.write(img_bytes)

            matched_df = match_photos_to_prices(photo_files, price_df)

            if matched_df.empty:
                st.error(
                    "No photos could be matched to codes in the Excel. "
                    "Please check that the photo filenames contain the correct product codes."
                )
                return

            try:
                pdf_bytes = build_pdf(matched_df, tmpdir)
            except Exception as e:
                st.error(f"Error while building PDF: {e}")
                return

            try:
                excel_bytes = build_excel(matched_df)
            except Exception as e:
                st.error(f"Error while building Excel: {e}")
                return

        st.success(
            f"Done! Matched {len(matched_df)} photos. "
            "Download your catalogue files below:"
        )

        st.download_button(
            "Download PDF catalogue",
            data=pdf_bytes,
            file_name="photo_catalogue.pdf",
            mime="application/pdf",
        )

        st.download_button(
            "Download Excel with matches",
            data=excel_bytes,
            file_name="photo_catalogue_matches.xlsx",
            mime=(
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            ),
        )


if __name__ == "__main__":
    main()
