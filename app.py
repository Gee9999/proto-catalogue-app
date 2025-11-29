import os
import re
import tempfile
from io import BytesIO

import pandas as pd
import streamlit as st
from fpdf import FPDF


# ---------- Helpers to detect columns in the price file ----------

def normalize_col(name: str) -> str:
    """Normalize a column header for matching: uppercase, remove non-alphanumerics."""
    return re.sub(r"[^A-Z0-9]+", "", str(name).upper())


def load_price_file(uploaded_file) -> pd.DataFrame:
    """
    Load the price Excel and return a clean DataFrame with columns:
      CODE, DESCRIPTION, PRICE_A_INCL, CODE_KEY
    CODE_KEY is a normalized numeric string used for matching filenames.
    """
    # Read everything as string so we don't lose leading zeros
    df = pd.read_excel(uploaded_file, dtype=str)

    # Find likely CODE, DESCRIPTION, PRICE columns
    code_col = None
    desc_col = None
    price_col = None

    for col in df.columns:
        norm = normalize_col(col)

        # CODE
        if code_col is None and norm in ("CODE", "ITEMCODE", "STOCKCODE", "PLUCODE"):
            code_col = col

        # DESCRIPTION
        if desc_col is None and (
            "DESCRIPTION" in norm
            or norm in ("DESC", "DESCR", "ITEMDESC", "STOCKDESC")
        ):
            desc_col = col

        # PRICE (PRICE-A INCL etc.)
        if price_col is None:
            if "PRICEAINCL" in norm or "PRICEAINCL" in norm:
                price_col = col
            elif norm.startswith("PRICEA") and "INCL" in norm:
                price_col = col

    if code_col is None:
        raise ValueError("Could not find CODE column in the price file.")
    if desc_col is None:
        raise ValueError("Could not find DESCRIPTION column in the price file.")
    if price_col is None:
        raise ValueError("Could not find price column (PRICE-A INCL).")

    df = df[[code_col, desc_col, price_col]].copy()
    df.columns = ["CODE", "DESCRIPTION", "PRICE_A_INCL"]

    # Drop rows with no code
    df["CODE"] = df["CODE"].astype(str).str.strip()
    df = df[df["CODE"].notna() & (df["CODE"] != "")]

    # Normalize price as string
    df["PRICE_A_INCL"] = df["PRICE_A_INCL"].astype(str).str.strip()

    # CODE_KEY: only digits from CODE, used to match filename digits
    df["CODE_KEY"] = (
        df["CODE"]
        .astype(str)
        .str.replace(r"[^0-9]", "", regex=True)
        .str.lstrip("0")  # remove leading zeros for safer matching
    )

    # Drop rows where CODE_KEY is empty
    df = df[df["CODE_KEY"] != ""]

    # Handle duplicates: keep the first row per CODE_KEY
    df = df.drop_duplicates(subset=["CODE_KEY"], keep="first")

    return df


# ---------- Match photos to price rows ----------

def extract_code_from_filename(filename: str) -> str | None:
    """
    Extract digits from filename like '8613900001-25PCS.JPG' -> '8613900001'.
    Returns normalized numeric string without leading zeros.
    """
    m = re.search(r"(\d+)", filename)
    if not m:
        return None
    digits = m.group(1)
    return digits.lstrip("0") or digits  # avoid empty if all zeros


def match_photos_to_prices(photo_files, price_df: pd.DataFrame) -> pd.DataFrame:
    """
    Given a list of uploaded photo files and a cleaned price_df,
    return a DataFrame with columns:

      PHOTO_FILE, CODE, DESCRIPTION, PRICE_A_INCL
    """
    # Build a lookup dict from CODE_KEY
    price_dict = (
        price_df.set_index("CODE_KEY")[["CODE", "DESCRIPTION", "PRICE_A_INCL"]]
        .to_dict("index")
    )

    rows = []

    for f in photo_files:
        fname = f.name
        key = extract_code_from_filename(fname)
        code_val = ""
        desc_val = ""
        price_val = ""

        if key:
            info = price_dict.get(key)
            if info is None:
                # Try without stripping zeros (just in case)
                info = price_dict.get(key.zfill(len(key)))
            if info is not None:
                code_val = info.get("CODE", "")
                desc_val = info.get("DESCRIPTION", "")
                price_val = info.get("PRICE_A_INCL", "")

        rows.append(
            {
                "PHOTO_FILE": fname,
                "CODE": code_val,
                "DESCRIPTION": desc_val,
                "PRICE_A_INCL": price_val,
            }
        )

    return pd.DataFrame(rows)


# ---------- PDF generation (60 x 60 images, text underneath) ----------

class CataloguePDF(FPDF):
    def __init__(self):
        super().__init__("P", "mm", "A4")
        self.set_auto_page_break(auto=True, margin=10)


def build_pdf(df: pd.DataFrame, photo_files, temp_dir: str) -> bytes:
    """
    Build a PDF with each product:
      - Photo (60 x 60 mm)
      - Code
      - Description
      - Price
    All text is *under* the photo.
    """
    pdf = CataloguePDF()
    pdf.add_page()
    pdf.set_font("Arial", "B", 14)
    pdf.cell(0, 10, "Photo Catalogue", ln=1, align="C")
    pdf.ln(2)

    # Map filename -> UploadedFile
    photo_map = {f.name: f for f in photo_files}

    # Layout: 2 columns per row
    page_w = 210
    margin_x = 10
    margin_top = 20
    usable_w = page_w - 2 * margin_x
    col_w = usable_w / 2.0

    img_size = 60  # <- 60 x 60 as requested
    text_height = 18  # rough height for 3 text lines
    row_height = img_size + text_height + 6

    x_positions = [margin_x, margin_x + col_w]
    y = margin_top
    col_index = 0

    pdf.set_font("Arial", size=9)

    for _, row in df.iterrows():
        # New page if we don't have enough space
        if y + row_height > (297 - 10):  # A4 height 297mm
            pdf.add_page()
            pdf.set_font("Arial", "B", 14)
            pdf.cell(0, 10, "Photo Catalogue", ln=1, align="C")
            pdf.ln(2)
            pdf.set_font("Arial", size=9)
            y = margin_top
            col_index = 0

        x = x_positions[col_index]

        # Draw image if available
        fname = row["PHOTO_FILE"]
        uploaded = photo_map.get(fname)
        if uploaded is not None:
            # Save to temp file and embed
            img_path = os.path.join(temp_dir, fname)
            with open(img_path, "wb") as img_out:
                img_out.write(uploaded.getvalue())
            pdf.image(img_path, x=x + (col_w - img_size) / 2, y=y, w=img_size, h=img_size)

        # Text underneath
        text_x = x
        text_y = y + img_size + 2
        pdf.set_xy(text_x, text_y)

        code = str(row.get("CODE", "") or "")
        desc = str(row.get("DESCRIPTION", "") or "")
        price = str(row.get("PRICE_A_INCL", "") or "")

        # Build a small block of text; multi_cell keeps it in the column
        lines = []
        if code:
            lines.append(f"Code: {code}")
        if desc:
            lines.append(desc)
        if price:
            lines.append(f"Price: {price}")
        text = "\n".join(lines) if lines else ""

        pdf.multi_cell(col_w, 4, text)

        # Move to next column / row
        if col_index == 0:
            col_index = 1
        else:
            col_index = 0
            y += row_height

    # fpdf2 dest="S" returns a bytearray; normalise to bytes
    result = pdf.output(dest="S")
    if isinstance(result, str):
        return result.encode("latin1")
    else:
        return bytes(result)


# ---------- Excel export ----------

def build_excel(df: pd.DataFrame) -> bytes:
    """
    Build an Excel file similar to the 'lovable perfect' one:

      Photo | Code | Description | Price

    Photo contains the filename of the image (so you can still see which is which).
    """
    export_df = pd.DataFrame(
        {
            "Photo": df["PHOTO_FILE"],
            "Code": df["CODE"],
            "Description": df["DESCRIPTION"],
            "Price": df["PRICE_A_INCL"],
        }
    )

    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        export_df.to_excel(writer, index=False, sheet_name="Catalogue")
    return output.getvalue()


# ---------- Streamlit app ----------

def main():
    st.set_page_config(page_title="Photo Catalogue Builder", layout="wide")
    st.title("Photo Catalogue Builder")

    st.markdown(
        """
        **Steps:**
        1. Upload your **price Excel** (any layout, as long as it has a CODE, DESCRIPTION and VAT-inclusive price column like `PRICE-A INCL.`).
        2. Upload the **product photos** (filenames must contain the product code digits, e.g. `8613900001-25PCS.JPG`).
        3. Click **Generate catalogue** to get:
           - A **PDF photo catalogue** (60×60 photo, with Code, Description, Price underneath).
           - An **Excel file** with columns **Photo, Code, Description, Price**.
        """
    )

    price_file = st.file_uploader(
        "Upload price Excel (XLS / XLSX)", type=["xls", "xlsx"], key="price_file"
    )

    photo_files = st.file_uploader(
        "Upload product photos (JPG / JPEG / PNG)",
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
            price_df = load_price_file(price_file)
        except Exception as e:
            st.error(f"Error reading price file: {e}")
            return

        try:
            df_matched = match_photos_to_prices(photo_files, price_df)
        except Exception as e:
            st.error(f"Error matching photos to prices: {e}")
            return

        if df_matched.empty:
            st.warning("No matches found between photo filenames and price codes.")
            return

        # Show a preview (no UploadedFile objects inside the DataFrame)
        st.subheader("Preview of matched data")
        st.dataframe(
            df_matched[["PHOTO_FILE", "CODE", "DESCRIPTION", "PRICE_A_INCL"]],
            use_container_width=True,
        )

        # Build files
        with tempfile.TemporaryDirectory() as tmpdir:
            try:
                pdf_bytes = build_pdf(df_matched, photo_files, tmpdir)
            except Exception as e:
                st.error(f"Error building PDF: {e}")
                return

        try:
            excel_bytes = build_excel(df_matched)
        except Exception as e:
            st.error(f"Error building Excel: {e}")
            return

        st.subheader("Download files")

        st.download_button(
            "Download PDF catalogue",
            data=pdf_bytes,
            file_name="photo_catalogue.pdf",
            mime="application/pdf",
        )

        st.download_button(
            "Download Excel file",
            data=excel_bytes,
            file_name="catalogue.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )


if __name__ == "__main__":
    main()
