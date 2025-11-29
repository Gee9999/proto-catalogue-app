import streamlit as st
import pandas as pd
import os
import io
import re
import tempfile
from fpdf import FPDF


# ---------- Helpers ----------

def norm_col(col_name: str) -> str:
    """
    Normalise a column name:
    - upper case
    - strip all non-alphanumeric characters
    """
    return re.sub(r"[^A-Z0-9]", "", str(col_name).upper())


def clean_code(val) -> str:
    """
    Take any value (from Excel or filename), extract digits only,
    and keep at most the first 10 digits as the code.
    """
    s = "".join(ch for ch in str(val) if ch.isdigit())
    return s[:10]


def load_price_file(uploaded_file) -> pd.DataFrame:
    """
    Load the price Excel and return a DataFrame with columns:
    CODE, DESCRIPTION, PRICE_A_INCL

    Looks for columns whose normalised names are:
    - CODE
    - DESCRIPTION
    - PRICEAINCL (for PRICE-A INCL. or similar)
    """
    filename = uploaded_file.name.lower()

    # Only Excel is needed here
    if filename.endswith((".xls", ".xlsx")):
        df = pd.read_excel(uploaded_file)
    else:
        raise ValueError("Unsupported price file type. Please upload an Excel file (.xls or .xlsx).")

    # Drop fully empty rows/cols
    df = df.dropna(how="all")
    df = df.loc[:, ~df.columns.to_series().apply(lambda c: df[c].isna().all())]

    # Map normalised column names
    norm_map = {col: norm_col(col) for col in df.columns}

    def find_col(target_norm: str):
        for col, n in norm_map.items():
            if n == target_norm:
                return col
        return None

    code_col = find_col("CODE")
    if not code_col:
        raise ValueError("Could not find CODE column.")

    desc_col = find_col("DESCRIPTION")
    if not desc_col:
        raise ValueError("Could not find DESCRIPTION column.")

    price_col = None
    for target in ["PRICEAINCL", "PRICEAINCLINC", "PRICEAINCLINCL"]:
        price_col = find_col(target)
        if price_col:
            break
    if not price_col:
        raise ValueError("Could not find price column (PRICE-A INCL.).")

    out = pd.DataFrame()
    out["CODE"] = df[code_col]
    out["DESCRIPTION"] = df[desc_col]

    price_series = df[price_col]

    # Clean price strings if needed
    if price_series.dtype == object:
        price_series = (
            price_series.astype(str)
            .str.replace("R", "", regex=False)
            .str.replace(",", "")
            .str.replace(" ", "")
        )

    out["PRICE_A_INCL"] = pd.to_numeric(price_series, errors="coerce")

    return out


def match_photos_to_prices(photo_files, price_df):
    """
    Match each uploaded photo to a row in price_df, based on the numeric code
    in the filename (first 10 digits).

    Returns:
      - df_excel: DataFrame with CODE, DESCRIPTION, PRICE_A_INCL
      - pdf_rows: list of dicts with PHOTO_FILE + text fields for PDF
    """
    # Work on a copy
    price_df = price_df.copy()

    # Remove rows without CODE
    price_df = price_df[price_df["CODE"].notna()]

    # Make a clean code string for lookup
    price_df["CODE_STR"] = price_df["CODE"].apply(clean_code)
    price_df = price_df[price_df["CODE_STR"] != ""]

    # FIX: remove duplicate codes so index is unique
    price_df = price_df.drop_duplicates(subset="CODE_STR", keep="first")

    # Build lookup dict
    price_dict = (
        price_df
        .set_index("CODE_STR")[["DESCRIPTION", "PRICE_A_INCL"]]
        .to_dict("index")
    )

    excel_rows = []
    pdf_rows = []

    for photo in photo_files:
        fname = photo.name
        code_str = clean_code(fname)

        if code_str in price_dict:
            info = price_dict[code_str]
            desc = info["DESCRIPTION"]
            price_val = info["PRICE_A_INCL"]
        else:
            desc = ""
            price_val = ""

        # For Excel (NO photo object)
        excel_rows.append({
            "CODE": code_str,
            "DESCRIPTION": desc,
            "PRICE_A_INCL": price_val,
        })

        # For PDF (needs photo object)
        pdf_rows.append({
            "PHOTO_FILE": photo,
            "CODE": code_str,
            "DESCRIPTION": desc,
            "PRICE_A_INCL": price_val,
        })

    df_excel = pd.DataFrame(excel_rows)
    return df_excel, pdf_rows


def build_pdf(pdf_rows, temp_dir: str) -> bytes:
    """
    Build the photo catalogue PDF.
    - 60x60 image
    - Code, description, price underneath the photo
    - 2 columns per row
    """
    pdf = FPDF(unit="mm", format="A4")
    pdf.set_auto_page_break(auto=False)
    pdf.add_page()

    # Title
    pdf.set_font("Arial", "B", 14)
    pdf.cell(0, 10, "Photo Catalogue", ln=1, align="C")
    pdf.ln(5)

    # Layout
    margin_x = 10
    margin_y = 15
    pdf.set_y(margin_y)

    page_w = pdf.w
    page_h = pdf.h

    photo_w = 60
    photo_h = 60
    text_line_h = 4
    extra_space = 6  # space below text
    block_h = photo_h + 3 * text_line_h + extra_space

    # Two columns
    col_gap = 10
    usable_w = page_w - 2 * margin_x
    col_w = (usable_w - col_gap) / 2

    x_left = margin_x
    x_right = margin_x + col_w + col_gap

    current_y = pdf.get_y()
    current_col = 0  # 0 = left, 1 = right

    pdf.set_font("Arial", size=8)

    for row in pdf_rows:
        # If not enough space on page for another block, start new page
        if current_y + block_h > page_h - margin_y:
            pdf.add_page()
            current_y = margin_y
            current_col = 0

        x = x_left if current_col == 0 else x_right

        # Save image to temp file
        photo = row["PHOTO_FILE"]
        img_path = os.path.join(temp_dir, photo.name)
        if not os.path.exists(img_path):
            with open(img_path, "wb") as f:
                f.write(photo.getvalue())

        # Draw image
        pdf.image(img_path, x=x, y=current_y, w=photo_w, h=photo_h)

        # Text below image
        text_y = current_y + photo_h + 2
        pdf.set_xy(x, text_y)

        code_text = str(row.get("CODE", "") or "")
        desc_text = str(row.get("DESCRIPTION", "") or "")
        price_val = row.get("PRICE_A_INCL", "")

        if price_val == "" or pd.isna(price_val):
            price_text = ""
        else:
            try:
                price_text = f"Price: {float(price_val):.2f}"
            except Exception:
                price_text = f"Price: {price_val}"

        # Multi-line text: CODE, DESCRIPTION, PRICE
        text_block = code_text
        if desc_text:
            text_block += "\n" + desc_text
        if price_text:
            text_block += "\n" + price_text

        pdf.multi_cell(photo_w, text_line_h, text_block, border=0)

        # Move to next column / row
        if current_col == 0:
            current_col = 1
        else:
            current_col = 0
            current_y = current_y + block_h

    # Return as bytes, avoid .encode() issue
    data = pdf.output(dest="S")
    if isinstance(data, (bytes, bytearray)):
        return bytes(data)
    return str(data).encode("latin1")


# ---------- Streamlit app ----------

def main():
    st.set_page_config(page_title="Photo Catalogue Builder", layout="centered")
    st.title("Photo Catalogue Builder")

    st.markdown(
        """
Upload your **price Excel** and **product photos**, and I'll build:

- A **PDF photo catalogue** (photos with code, description & price underneath).
- An **Excel file** with columns: `CODE`, `DESCRIPTION`, `PRICE_A_INCL`.
        """
    )

    price_file = st.file_uploader(
        "Upload price Excel file (e.g. PRODUCT DETAILS - BY CODE.xlsx)",
        type=["xls", "xlsx"],
        key="price_file",
    )

    photo_files = st.file_uploader(
        "Upload product photos (filenames must contain the product code, e.g. 8613900001-25pcs.jpg)",
        type=["jpg", "jpeg", "png"],
        accept_multiple_files=True,
        key="photo_files",
    )

    if st.button("Generate catalogue"):
        if not price_file:
            st.error("Please upload a price Excel file.")
            return
        if not photo_files:
            st.error("Please upload at least one product photo.")
            return

        try:
            with st.spinner("Processing prices..."):
                price_df = load_price_file(price_file)

            with st.spinner("Matching photos to prices..."):
                df_excel, pdf_rows = match_photos_to_prices(photo_files, price_df)

            # Show preview of Excel data (safe: no UploadedFile objects)
            st.subheader("Matched items preview")
            st.dataframe(df_excel)

            # Build PDF in a temp dir
            with st.spinner("Building PDF catalogue..."):
                with tempfile.TemporaryDirectory() as temp_dir:
                    pdf_bytes = build_pdf(pdf_rows, temp_dir)

            st.subheader("Download files")

            # PDF download
            st.download_button(
                "Download photo catalogue (PDF)",
                data=pdf_bytes,
                file_name="photo_catalogue.pdf",
                mime="application/pdf",
            )

            # Excel download (only CODE, DESCRIPTION, PRICE_A_INCL)
            out = io.BytesIO()
            df_excel.to_excel(out, index=False)
            out.seek(0)

            st.download_button(
                "Download Excel with CODE, DESCRIPTION & PRICE",
                data=out.getvalue(),
                file_name="catalogue.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

        except Exception as e:
            st.error(f"Something went wrong: {e}")


if __name__ == "__main__":
    main()
