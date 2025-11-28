import io
import os
import re
from typing import Dict, List, Tuple

import pandas as pd
import streamlit as st
from PIL import Image
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib.utils import ImageReader
from reportlab.pdfgen import canvas


# ---------- Helpers for Excel ----------

def normalise_colname(col: str) -> str:
    if not isinstance(col, str):
        col = str(col)
    return (
        col.upper()
        .replace(" ", "")
        .replace("\n", "")
        .replace("\r", "")
        .replace("_X000A_", "")
    )


def detect_columns(df: pd.DataFrame) -> Tuple[str, str, str]:
    """
    Try to guess the CODE, DESCRIPTION and PRICE-A INCL columns.
    Returns (code_col, desc_col, price_col).
    Raises ValueError if something important is missing.
    """
    code_col = None
    desc_col = None
    price_col = None

    for col in df.columns:
        n = normalise_colname(col)
        if code_col is None and "CODE" in n:
            code_col = col
        if desc_col is None and ("DESC" in n or "DESCRIPTION" in n):
            desc_col = col
        if price_col is None and "PRICE" in n and "INCL" in n:
            price_col = col

    if code_col is None:
        raise ValueError("Could not find CODE column in the price file.")
    if desc_col is None:
        raise ValueError("Could not find DESCRIPTION column in the price file.")
    if price_col is None:
        raise ValueError("Could not find VAT inclusive price column in the price file.")

    return code_col, desc_col, price_col


def load_price_table(excel_file) -> pd.DataFrame:
    df = pd.read_excel(excel_file)
    code_col, desc_col, price_col = detect_columns(df)

    # Clean up / standardise
    df = df[[code_col, desc_col, price_col]].copy()
    df.columns = ["CODE", "DESCRIPTION", "PRICE_A_INCL"]

    # Normalise code to a clean numeric string
    df["CODE_STR"] = (
        df["CODE"]
        .astype(str)
        .str.strip()
        .str.extract(r"(\d+)", expand=False)
    )

    return df


# ---------- Helpers for photos & matching ----------

def extract_code_from_filename(filename: str) -> str:
    """
    Extract the longest leading number sequence from the filename.
    e.g. '8613900001-25pcs.jpg' -> '8613900001'
    """
    base = os.path.basename(filename)
    # Take part before first non-digit after starting digits
    m = re.match(r"(\d+)", base)
    if m:
        return m.group(1)
    # Fallback: first digit group anywhere
    m = re.search(r"(\d+)", base)
    return m.group(1) if m else ""


def match_photos_to_prices(
    photos: List[st.runtime.uploaded_file_manager.UploadedFile],
    price_df: pd.DataFrame,
) -> pd.DataFrame:
    """
    Exact matching: photo code must match CODE_STR.
    Returns a DataFrame with one row per photo that found a match.
    """
    code_index: Dict[str, Dict[str, object]] = (
        price_df.set_index("CODE_STR")[["CODE", "DESCRIPTION", "PRICE_A_INCL"]].to_dict("index")
    )

    records = []
    for up in photos:
        fname = up.name
        code = extract_code_from_filename(fname)
        info = code_index.get(code)
        if info is None:
            # No match – just skip this photo
            continue

        records.append(
            {
                "PHOTO_FILE": fname,
                "CODE": info["CODE"],
                "DESCRIPTION": info["DESCRIPTION"],
                "PRICE-A INCL.": info["PRICE_A_INCL"],
            }
        )

    return pd.DataFrame(records)


# ---------- PDF generation (no encode, returns BytesIO) ----------

def build_pdf_catalog(matched_df: pd.DataFrame, photos: List[st.runtime.uploaded_file_manager.UploadedFile]) -> io.BytesIO:
    """
    Build a simple photo catalogue PDF with 2 columns:
    image + CODE + DESCRIPTION + PRICE-A INCL.
    """
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)
    page_w, page_h = A4

    margin = 15 * mm
    col_gap = 5 * mm
    num_cols = 2
    cell_w = (page_w - 2 * margin - col_gap * (num_cols - 1)) / num_cols
    max_img_h = 40 * mm
    text_line_height = 9

    # Map photo name -> UploadedFile for easy lookup
    photo_map = {p.name: p for p in photos}

    y = page_h - margin
    col = 0

    # Ensure column name is easy to use
    df = matched_df.copy()
    if "PRICE-A INCL." in df.columns:
        df = df.rename(columns={"PRICE-A INCL.": "PRICE"})

    for _, row in df.iterrows():
        photo_name = row["PHOTO_FILE"]
        up = photo_map.get(photo_name)
        if up is None:
            continue

        # Convert uploaded bytes to an ImageReader
        img_bytes = io.BytesIO(up.getvalue())
        try:
            img_reader = ImageReader(img_bytes)
            iw, ih = img_reader.getSize()
        except Exception:
            # If image is somehow unreadable, skip image but still print text
            img_reader = None
            iw, ih = (1, 1)

        scale = min((cell_w - 10) / iw, max_img_h / ih) if img_reader else 1.0
        img_w = iw * scale
        img_h = ih * scale if img_reader else 0

        # Start new row if needed
        needed_height = img_h + 3 * text_line_height + 15
        if y - needed_height < margin:
            c.showPage()
            y = page_h - margin
            col = 0

        if col == 0:
            x = margin
        else:
            x = margin + cell_w + col_gap

        # Draw image
        if img_reader:
            img_x = x + (cell_w - img_w) / 2
            img_y = y - img_h
            c.drawImage(img_reader, img_x, img_y, width=img_w, height=img_h, preserveAspectRatio=True)
        else:
            img_y = y

        text_y = img_y - 8
        c.setFont("Helvetica-Bold", 8)
        c.drawString(x, text_y, f"CODE: {row['CODE']}")
        text_y -= text_line_height

        c.setFont("Helvetica", 8)
        desc = str(row.get("DESCRIPTION", "") or "")
        # Simple truncation – you can improve with manual wrapping later
        c.drawString(x, text_y, f"DESC: {desc[:70]}")
        text_y -= text_line_height

        price_val = row.get("PRICE", "")
        c.drawString(x, text_y, f"PRICE (incl): {price_val}")
        text_y -= text_line_height

        # Move to next column / row
        if col == 0:
            col = 1
        else:
            col = 0
            y = text_y - 10

    c.save()
    buf.seek(0)
    return buf


# ---------- Excel generation ----------

def build_excel_catalog(matched_df: pd.DataFrame) -> io.BytesIO:
    """
    Create a simple Excel with columns:
    PHOTO_FILE, CODE, DESCRIPTION, PRICE-A INCL.
    (and NO images – just data)
    """
    buf = io.BytesIO()
    out_df = matched_df[["PHOTO_FILE", "CODE", "DESCRIPTION", "PRICE-A INCL."]].copy()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        out_df.to_excel(writer, index=False, sheet_name="Catalogue")
    buf.seek(0)
    return buf


# ---------- Streamlit UI ----------

def main():
    st.title("Photo Catalogue Builder")

    st.markdown(
        """
        1. Upload your **price Excel file** (any layout, as long as it has a CODE, DESCRIPTION and VAT-inclusive price column).  
        2. Upload the **product photos** (filenames must contain the product code, e.g. `8613900001-25pcs.jpg`).  
        3. Click **Generate catalogue** to get:
           - A **PDF photo catalogue**  
           - A matching **Excel file** with columns `PHOTO_FILE`, `CODE`, `DESCRIPTION`, `PRICE-A INCL.`
        """
    )

    price_file = st.file_uploader(
        "Upload price Excel (e.g. PRODUCT DETAILS - BY CODE.xlsx)",
        type=["xls", "xlsx"],
        key="price_file",
    )

    photo_files = st.file_uploader(
        "Upload product photos",
        type=["jpg", "jpeg", "png"],
        accept_multiple_files=True,
        key="photo_files",
    )

    if st.button("Generate catalogue"):

        if not price_file:
            st.error("Please upload a price Excel file first.")
            return
        if not photo_files:
            st.error("Please upload at least one product photo.")
            return

        try:
            price_df = load_price_table(price_file)
        except Exception as e:
            st.error(f"Couldn't read price file: {e}")
            return

        matched_df = match_photos_to_prices(photo_files, price_df)

        if matched_df.empty:
            st.warning("No photos matched any product codes from the Excel file.")
            return

        st.success(f"Matched {len(matched_df)} photos to products.")
        st.dataframe(matched_df)

        # Build Excel
        excel_buf = build_excel_catalog(matched_df)

        # Build PDF
        pdf_buf = build_pdf_catalog(matched_df, photo_files)

        st.download_button(
            "Download Excel catalogue",
            data=excel_buf,
            file_name="photo_catalogue.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        st.download_button(
            "Download PDF catalogue",
            data=pdf_buf,
            file_name="photo_catalogue.pdf",
            mime="application/pdf",
        )


if __name__ == "__main__":
    main()
