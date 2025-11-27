import streamlit as st
import pandas as pd
import fitz  # PyMuPDF
import re
from io import BytesIO
from PIL import Image as PILImage

from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage

from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader


st.set_page_config(page_title="Proto Catalogue Builder", layout="wide")
st.title("📸 Proto Trading – Product Catalogue Builder")


# -------------------------
# Helpers
# -------------------------
def normalize_code(code: str) -> str:
    """Keep only digits: '8610100003N' -> '8610100003'."""
    return re.sub(r"[^0-9]", "", str(code))


# -------------------------
# 1) FAST PRICE EXTRACTION FROM PDF
# -------------------------
@st.cache_data
def extract_prices_fast(pdf_bytes: bytes) -> pd.DataFrame:
    """
    Read the big 'PRODUCT DETAILS - BY CODE.pdf' and extract:
    CODE, DESCRIPTION, PRICE-A INCL
    """
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")

    # CODE at line start, digits + optional letter
    code_pattern = re.compile(r"^(\d+[A-Za-z]?)\b")

    price_dict = {}  # keyed by normalized digits

    for page in doc:
        text = page.get_text("text")
        if not text:
            continue

        for line in text.split("\n"):
            line = line.strip()
            if not line:
                continue

            m = code_pattern.match(line)
            if not m:
                continue

            code_raw = m.group(1)  # e.g. 8610100003N
            norm_code = normalize_code(code_raw)

            # Already have this code? Keep the first occurrence
            if norm_code in price_dict:
                continue

            # All decimal numbers in the row
            numbers = re.findall(r"\d+\.\d+", line)
            if len(numbers) < 5:
                # We expect at least 5 decimals; 5th = PRICE-A INCL
                continue

            try:
                price_incl = float(numbers[4])
            except ValueError:
                continue

            # DESCRIPTION: words after code until first decimal
            parts = line.split()
            desc_tokens = []
            for p in parts[1:]:
                if re.match(r"\d+\.\d+", p):
                    break
                desc_tokens.append(p)
            description = " ".join(desc_tokens)

            price_dict[norm_code] = {
                "CODE": code_raw,
                "NORM_CODE": norm_code,
                "DESCRIPTION": description,
                "PRICE_INCL": price_incl,
            }

    df = pd.DataFrame(price_dict.values())
    return df


# -------------------------
# 2) MATCH PHOTOS TO PRICES
# -------------------------
def extract_photo_norm_code(filename: str) -> str:
    """
    From filename like '86101000001-10mm-10pcs.jpg'
    -> take part before '-', keep only digits.
    """
    stem = filename.rsplit(".", 1)[0]
    base = stem.split("-")[0]
    return normalize_code(base)


def match_photos_to_prices(photo_files, price_df: pd.DataFrame) -> pd.DataFrame:
    """
    For each photo:
      - compute a normalized numeric code from filename
      - try to find a matching row in price_df by NORM_CODE
      - if not found, try dropping last digit(s) as a fallback
    """
    # Build a dict for fast lookup by normalized code
    price_map = {}
    for _, row in price_df.iterrows():
        norm = str(row["NORM_CODE"])
        price_map[norm] = row

    rows = []

    for file in photo_files:
        filename = file.name
        photo_norm = extract_photo_norm_code(filename)

        code_raw = ""
        desc = ""
        price_val = None

        row = None
        if photo_norm:
            # 1) Exact match on normalized digits
            row = price_map.get(photo_norm)

            # 2) If no exact match, try trimming last 1–3 digits
            if row is None:
                for trim in range(1, 4):
                    if len(photo_norm) - trim < 4:
                        break
                    candidate = photo_norm[:-trim]
                    row = price_map.get(candidate)
                    if row is not None:
                        break

        if row is not None:
            code_raw = row["CODE"]
            desc = row["DESCRIPTION"]
            price_val = row["PRICE_INCL"]
        else:
            # No match found – keep the numeric code we parsed at least
            code_raw = photo_norm
            desc = ""
            price_val = None

        # Load image once here
        img = PILImage.open(BytesIO(file.getvalue())).convert("RGB")

        rows.append(
            {
                "IMAGE": img,
                "CODE": code_raw,
                "DESCRIPTION": desc,
                "PRICE_INCL": price_val,
                "FILENAME": filename,
            }
        )

    return pd.DataFrame(rows)


# -------------------------
# 3) EXCEL WITH THUMBNAILS
# -------------------------
def build_excel_with_thumbnails(matched_df: pd.DataFrame) -> BytesIO:
    """
    Excel layout:
      Col A: Photo thumbnail
      Col B: Code
      Col C: Description
      Col D: Price incl
      Col E: Filename
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "Products"

    headers = ["Photo", "Code", "Description", "Price incl", "Filename"]
    ws.append(headers)

    ws.column_dimensions["A"].width = 30
    ws.column_dimensions["B"].width = 18
    ws.column_dimensions["C"].width = 40
    ws.column_dimensions["D"].width = 14
    ws.column_dimensions["E"].width = 30

    row_idx = 2
    for _, row in matched_df.iterrows():
        code = row["CODE"]
        desc = row["DESCRIPTION"]
        price = row["PRICE_INCL"]
        filename = row["FILENAME"]
        img = row["IMAGE"]

        # Cells (text)
        ws.cell(row=row_idx, column=2, value=code)
        ws.cell(row=row_idx, column=3, value=desc)
        ws.cell(row=row_idx, column=5, value=filename)

        if price is not None and not pd.isna(price):
            try:
                ws.cell(row=row_idx, column=4, value=float(price))
            except ValueError:
                pass

        # Image thumbnail in col A
        try:
            thumb = img.copy()
            thumb.thumbnail((180, 180))
            img_buf = BytesIO()
            thumb.save(img_buf, format="PNG")
            img_buf.seek(0)

            xl_img = XLImage(img_buf)
            ws.add_image(xl_img, f"A{row_idx}")
            ws.row_dimensions[row_idx].height = 140
        except Exception as e:
            print("Excel image error:", e)

        row_idx += 1

    out = BytesIO()
    wb.save(out)
    out.seek(0)
    return out


# -------------------------
# 4) 3×3 PDF CATALOGUE – IMAGE, DESCRIPTION, PRICE, CODE
# -------------------------
def build_pdf_catalog(matched_df: pd.DataFrame) -> BytesIO:
    """
    3×3 grid per A4 page:
      [IMAGE]
      Description
      Price
      Code
    """
    buf = BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)
    page_w, page_h = A4

    margin_left = 40
    margin_right = 40
    margin_top = 50
    margin_bottom = 50

    cols = 3
    rows = 3
    usable_w = page_w - margin_left - margin_right
    usable_h = page_h - margin_top - margin_bottom - 30

    cell_w = usable_w / cols
    cell_h = usable_h / rows

    img_max_w = cell_w - 20
    img_max_h = cell_h * 0.55

    items_per_page = cols * rows

    for idx, (_, row) in enumerate(matched_df.iterrows()):
        pos = idx % items_per_page

        if pos == 0 and idx != 0:
            c.showPage()

        col = pos % cols
        r = pos // cols

        x0 = margin_left + col * cell_w
        y0 = margin_bottom + (rows - 1 - r) * cell_h

        img = row["IMAGE"]
        desc = str(row["DESCRIPTION"]) if row["DESCRIPTION"] else ""
        price = row["PRICE_INCL"]
        code_val = str(row["CODE"]) if row["CODE"] else ""

        if isinstance(price, (int, float)) and not pd.isna(price):
            price_str = f"R{price:,.2f}"  # format R42.50
        else:
            price_str = "R0.00" if desc or code_val else ""

        # Draw image
        img_height_used = 0
        try:
            pil_img = img.copy()
            pil_img.thumbnail((img_max_w, img_max_h))
            img_buf = BytesIO()
            pil_img.save(img_buf, format="PNG")
            img_buf.seek(0)

            iw, ih = pil_img.size
            img_reader = ImageReader(img_buf)
            img_x = x0 + (cell_w - iw) / 2
            img_y = y0 + cell_h - ih - 20

            c.drawImage(img_reader, img_x, img_y, width=iw, height=ih)
            img_height_used = ih
        except Exception as e:
            print("PDF image error:", e)

        # Text below image
        text_y = y0 + cell_h - img_height_used - 30

        # Description
        c.setFont("Helvetica", 8)
        c.drawString(x0 + 10, text_y, desc[:80])

        # Price
        c.setFont("Helvetica", 9)
        c.drawString(x0 + 10, text_y - 14, f"Price: {price_str}")

        # Code (barcode number) – LAST LINE
        c.setFont("Helvetica", 9)
        c.drawString(x0 + 10, text_y - 28, f"Code: {code_val}")

    c.showPage()
    c.save()
    buf.seek(0)
    return buf


# -------------------------
# STREAMLIT UI
# -------------------------
st.header("1️⃣ Upload price PDF (PRODUCT DETAILS - BY CODE.pdf)")
price_pdf = st.file_uploader("Price PDF", type=["pdf"])

st.header("2️⃣ Upload product photos")
photos = st.file_uploader(
    "Product photos (filenames must contain the code)",
    accept_multiple_files=True,
    type=["jpg", "jpeg", "png"],
)

if st.button("PROCESS"):
    if not price_pdf or not photos:
        st.error("Please upload BOTH the price PDF and at least one photo.")
    else:
        try:
            with st.spinner("Extracting prices from PDF (fast)..."):
                pdf_bytes = price_pdf.read()
                price_df = extract_prices_fast(pdf_bytes)

            st.success(f"Extracted {len(price_df)} items from price PDF.")

            with st.spinner("Matching photos to codes and prices..."):
                matched_df = match_photos_to_prices(photos, price_df)

            # For Streamlit display, drop IMAGE column
            display_df = matched_df.drop(columns=["IMAGE"])
            st.subheader("Matched data preview")
            st.dataframe(display_df)

            with st.spinner("Building Excel with thumbnails..."):
                excel_file = build_excel_with_thumbnails(matched_df)

            with st.spinner("Building 3×3 PDF catalogue..."):
                pdf_file = build_pdf_catalog(matched_df)

            st.success("Done! Download your files below:")

            st.download_button(
                "⬇️ Download Excel",
                data=excel_file,
                file_name="product_catalogue.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

            st.download_button(
                "📄 Download 3×3 PDF Catalogue",
                data=pdf_file,
                file_name="product_catalogue.pdf",
                mime="application/pdf",
            )

        except Exception as e:
            st.error(f"Error while processing: {e}")
