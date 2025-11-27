import streamlit as st
import pandas as pd
import os
import re
import pdfplumber
from PIL import Image
from io import BytesIO
from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage
from fpdf import FPDF
import tempfile

st.set_page_config(page_title="Proto Catalogue (PDF only)", layout="wide")
st.title("📸 Proto Trading – Catalogue Builder (3×3 / 4×4, PDF Only)")


# ---------------------------------------------------------
# Helpers
# ---------------------------------------------------------
def normalize_code(code: str) -> str:
    """Keep only digits."""
    return re.sub(r"[^0-9]", "", str(code))


def extract_photo_norm_code(filename: str) -> str:
    """
    Extract numeric code from photo name:
    Example: 8613900012-20PCS.jpg → 8613900012
    """
    stem = os.path.splitext(filename)[0]
    base = stem.split("-")[0]
    return normalize_code(base)


# ---------------------------------------------------------
# Extract prices from PDF using pdfplumber (no OCR)
# ---------------------------------------------------------
@st.cache_data
def extract_prices_from_pdf(pdf_bytes: bytes) -> pd.DataFrame:
    """
    Extract CODE, DESCRIPTION, PRICE_INCL from 'PRODUCT DETAILS - BY CODE.pdf'
    Uses text-only extraction with pdfplumber.
    Assumes:
      - line starts with code (e.g. 8613900012 or 8613900012N)
      - there are at least 5 decimal numbers on the line
      - 5th decimal is PRICE-A INCL
    """
    items = {}
    code_pattern = re.compile(r"^(\d+[A-Za-z]?)\b")

    with pdfplumber.open(BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if not text:
                continue

            for raw_line in text.split("\n"):
                line = raw_line.strip()
                if not line:
                    continue

                m = code_pattern.match(line)
                if not m:
                    continue

                code_raw = m.group(1)
                norm = normalize_code(code_raw)

                # avoid duplicates, keep first occurrence
                if norm in items:
                    continue

                numbers = re.findall(r"\d+\.\d+", line)
                if len(numbers) < 5:
                    continue

                try:
                    price_incl = float(numbers[4])
                except ValueError:
                    continue

                parts = line.split()
                desc_tokens = []
                for p in parts[1:]:
                    if re.match(r"\d+\.\d+", p):
                        break
                    desc_tokens.append(p)

                description = " ".join(desc_tokens)

                items[norm] = {
                    "CODE": code_raw,
                    "NORM_CODE": norm,
                    "DESCRIPTION": description,
                    "PRICE_INCL": price_incl,
                }

    df = pd.DataFrame(items.values())
    return df


# ---------------------------------------------------------
# Match photos to prices
# ---------------------------------------------------------
def match_photos_to_prices(photo_names, price_df: pd.DataFrame) -> pd.DataFrame:
    """
    For each photo filename:
      - derive a normalized numeric code
      - try exact match on NORM_CODE
      - if not found, try trimming last 1–3 digits
    """
    price_map = {str(r["NORM_CODE"]): r for _, r in price_df.iterrows()}

    rows = []
    for fname in photo_names:
        norm = extract_photo_norm_code(fname)

        row = price_map.get(norm)

        # Fallback: trim last 1–3 digits
        if row is None and norm:
            for trim in range(1, 4):
                if len(norm) - trim < 4:
                    break
                candidate = norm[:-trim]
                row = price_map.get(candidate)
                if row is not None:
                    break

        if row is not None:
            rows.append({
                "PHOTO": fname,
                "CODE": row["CODE"],
                "DESCRIPTION": row["DESCRIPTION"],
                "PRICE_INCL": row["PRICE_INCL"],
                "FILENAME": fname
            })
        else:
            rows.append({
                "PHOTO": fname,
                "CODE": norm,
                "DESCRIPTION": "",
                "PRICE_INCL": None,
                "FILENAME": fname
            })

    return pd.DataFrame(rows)


# ---------------------------------------------------------
# Build Excel with thumbnail images (A:E)
# ---------------------------------------------------------
def build_excel_with_thumbnails(df: pd.DataFrame, photo_folder: str) -> BytesIO:
    """
    Excel format:
      Col A: Photo thumbnail
      Col B: Code
      Col C: Description
      Col D: Price incl
      Col E: Filename
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "Products"

    ws.append(["Photo", "Code", "Description", "Price incl", "Filename"])

    ws.column_dimensions["A"].width = 30
    ws.column_dimensions["B"].width = 18
    ws.column_dimensions["C"].width = 40
    ws.column_dimensions["D"].width = 14
    ws.column_dimensions["E"].width = 30

    row_idx = 2
    for _, row in df.iterrows():
        ws.cell(row=row_idx, column=2, value=row["CODE"])
        ws.cell(row=row_idx, column=3, value=row["DESCRIPTION"])
        ws.cell(row=row_idx, column=5, value=row["FILENAME"])

        price = row["PRICE_INCL"]
        if isinstance(price, (int, float)) and not pd.isna(price):
            ws.cell(row=row_idx, column=4, value=float(price))

        img_path = os.path.join(photo_folder, row["PHOTO"])
        try:
            xl_img = XLImage(img_path)
            xl_img.width = 130
            xl_img.height = 130
            ws.add_image(xl_img, f"A{row_idx}")
            ws.row_dimensions[row_idx].height = 110
        except Exception as e:
            print("Excel image error:", e)

        row_idx += 1

    out = BytesIO()
    wb.save(out)
    out.seek(0)
    return out


# ---------------------------------------------------------
# Build PDF catalogue (3×3 or 4×4)
# ---------------------------------------------------------
def generate_pdf_layout(df: pd.DataFrame, photo_folder: str, layout: str) -> bytes:
    """
    Layout options:
      - "3x3" → 3 columns x 3 rows
      - "4x4" → 4 columns x 4 rows

    Each cell:
      [IMAGE ~60mm wide]
      Description
      Price: Rxx.xx
      Code: 8613...
    """
    pdf = FPDF("P", "mm", "A4")
    pdf.set_auto_page_break(auto=True, margin=10)

    if layout == "3x3":
        cols, rows = 3, 3
    else:
        cols, rows = 4, 4

    page_w, page_h = 210, 297
    margin_x, margin_y = 10, 10
    usable_w = page_w - 2 * margin_x
    usable_h = page_h - 2 * margin_y

    cell_w = usable_w / cols
    cell_h = usable_h / rows

    img_w = 60
    img_h = 45  # approximate, keeps aspect ok

    items_per_page = cols * rows

    for i in range(0, len(df), items_per_page):
        pdf.add_page()
        chunk = df.iloc[i:i + items_per_page]

        for idx, (_, row) in enumerate(chunk.iterrows()):
            r = idx // cols
            c = idx % cols

            x = margin_x + c * cell_w
            y = margin_y + r * cell_h

            img_path = os.path.join(photo_folder, row["PHOTO"])
            desc = str(row["DESCRIPTION"]) if row["DESCRIPTION"] else ""
            code_val = str(row["CODE"]) if row["CODE"] else ""
            price_val = row["PRICE_INCL"]

            if isinstance(price_val, (int, float)) and not pd.isna(price_val):
                price_str = f"R{price_val:,.2f}"
            else:
                price_str = "R0.00" if desc or code_val else ""

            # Image centered in the cell
            try:
                pdf.image(img_path, x=x + (cell_w - img_w) / 2, y=y, w=img_w)
            except Exception as e:
                print("PDF image error:", e)

            text_y = y + img_h + 4

            pdf.set_xy(x, text_y)
            pdf.set_font("Arial", size=8)
            pdf.multi_cell(cell_w, 4, desc[:160], 0, "L")

            pdf.set_xy(x, text_y + 14)
            pdf.set_font("Arial", size=9)
            pdf.cell(cell_w, 4, f"Price: {price_str}", 0, 2, "L")

            pdf.set_xy(x, text_y + 21)
            pdf.cell(cell_w, 4, f"Code: {code_val}", 0, 2, "L")

    out = BytesIO()
    pdf.output(out)
    return out.getvalue()


# ---------------------------------------------------------
# Streamlit UI
# ---------------------------------------------------------
st.header("1️⃣ Upload Price PDF (PRODUCT DETAILS - BY CODE.pdf)")
price_pdf = st.file_uploader("Price PDF", type=["pdf"])

st.header("2️⃣ Upload Product Photos")
photos = st.file_uploader(
    "Product photos (filenames must contain the code)",
    type=["jpg", "jpeg", "png"],
    accept_multiple_files=True,
)

layout_choice = st.radio(
    "Choose PDF layout",
    ["3×3 grid", "4×4 grid"],
    index=0,
)

if st.button("PROCESS"):
    if not price_pdf or not photos:
        st.error("Please upload BOTH the price PDF and at least one photo.")
    else:
        with st.spinner("Extracting prices from PDF (text-only)…"):
            pdf_bytes = price_pdf.read()
            price_df = extract_prices_from_pdf(pdf_bytes)

        st.success(f"Extracted {len(price_df)} price rows from PDF.")

        with st.spinner("Saving photos and matching prices…"):
            temp_dir = tempfile.mkdtemp(prefix="proto_photos_")
            photo_names = []
            for pf in photos:
                photo_names.append(pf.name)
                save_path = os.path.join(temp_dir, pf.name)
                with open(save_path, "wb") as f:
                    f.write(pf.getvalue())

            matched_df = match_photos_to_prices(photo_names, price_df)

        st.subheader("Matched Results (Preview)")
        st.dataframe(matched_df)

        with st.spinner("Building Excel with thumbnails…"):
            excel_file = build_excel_with_thumbnails(matched_df, temp_dir)

        layout_key = "3x3" if "3×3" in layout_choice else "4x4"
        with st.spinner(f"Building {layout_key} PDF catalogue…"):
            pdf_file = generate_pdf_layout(matched_df, temp_dir, layout_key)

        st.success("Done! Download your files below:")

        st.download_button(
            "⬇️ Download Excel",
            data=excel_file,
            file_name="product_catalogue.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        st.download_button(
            f"📄 Download {layout_choice} PDF Catalogue",
            data=pdf_file,
            file_name="product_catalogue.pdf",
            mime="application/pdf",
        )
