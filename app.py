import streamlit as st
import pandas as pd
import os
import re
import pdfplumber
from PIL import Image
from io import BytesIO
import tempfile

from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage
from fpdf import FPDF

# ---------------------------------------------------------
# Streamlit config
# ---------------------------------------------------------
st.set_page_config(page_title="Proto Catalogue Builder", layout="wide")
st.title("📸 Proto Trading – Photo & Price Catalogue Builder")


# ---------------------------------------------------------
# Helpers
# ---------------------------------------------------------
def normalize_code(code: str) -> str:
    """Keep only digits from a code."""
    return re.sub(r"[^0-9]", "", str(code))


def extract_photo_norm_code(filename: str) -> str:
    """
    Extract numeric code from photo filename.

    Examples:
        8613900012-20PCS.jpg        -> 8613900012
        86101000001-10mm-10pcs.jpg  -> 86101000001
        86101000004A.jpg            -> 86101000004
    """
    stem = os.path.splitext(filename)[0]
    base = stem.split("-")[0]  # everything before first '-'
    return normalize_code(base)


# ---------------------------------------------------------
# PDF price extraction (block-based, no OCR)
# ---------------------------------------------------------
@st.cache_data
def extract_prices_from_pdf(pdf_bytes: bytes) -> pd.DataFrame:
    """
    Extract CODE, DESCRIPTION, PRICE_INCL from 'PRODUCT DETAILS - BY CODE.pdf'.

    Your sample block:

        8613900033
        FIMO BEAD STRING 3mm
         4.22
         4.22
         4.85
         12.61
         14.50 T
        N
         1.00
         43.00
         0.00
         0.00 Y
        N

    Rules used:
      - First line of a block = CODE (digits, maybe with trailing letter)
      - Second (and possibly more) text lines = DESCRIPTION
      - PRICE-A INCL = the decimal number immediately followed by ' T'
           e.g. "14.50 T"  -> 14.50
    """
    items = {}
    current_block = []

    def flush_block(block_lines):
        """Process one product block into items dict."""
        if not block_lines:
            return

        # 1) CODE line = first line, must start with digits
        first = block_lines[0].strip()
        if not re.match(r"^\d+[A-Za-z]?$", first):
            return

        code_raw = first
        norm_code = normalize_code(code_raw)
        if not norm_code:
            return

        # Avoid duplicates (keep first occurrence)
        if norm_code in items:
            return

        # 2) DESCRIPTION = consecutive non-numeric lines after code,
        #    until we hit a purely numeric/flag line
        desc_lines = []
        for line in block_lines[1:]:
            s = line.strip()
            if not s:
                continue
            # numeric or price-like / flags => stop description
            if re.match(r"^[0-9]+\.[0-9]+(\s+[A-Z])?$", s) or re.match(r"^[0-9]+\s*$", s) or s in ("T", "N", "Y"):
                break
            desc_lines.append(s)

        description = " ".join(desc_lines)

        # 3) PRICE_INCL = first decimal number followed by ' T'
        price_incl = None
        for line in block_lines:
            s = line.strip()
            m = re.search(r"(\d+\.\d+)\s+T\b", s)
            if m:
                try:
                    price_incl = float(m.group(1))
                    break
                except ValueError:
                    continue

        # If no price, we still keep code + description
        items[norm_code] = {
            "CODE": code_raw,
            "NORM_CODE": norm_code,
            "DESCRIPTION": description,
            "PRICE_INCL": price_incl,
        }

    # --- Read PDF and split into logical blocks ---
    with pdfplumber.open(BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if not text:
                continue

            lines = text.split("\n")
            for raw_line in lines:
                line = raw_line.rstrip("\n")
                stripped = line.strip()

                # Start of a new block = line beginning with a long code
                if re.match(r"^\d{6,}[A-Za-z]?$", stripped):
                    # flush previous block
                    flush_block(current_block)
                    current_block = [stripped]
                else:
                    # continue current block
                    if current_block:
                        current_block.append(line)

    # Flush last block
    flush_block(current_block)

    df = pd.DataFrame(items.values())
    return df


# ---------------------------------------------------------
# Match photos to price list
# ---------------------------------------------------------
def match_photos_to_prices(photo_names, price_df: pd.DataFrame) -> pd.DataFrame:
    """
    For each photo:
      - derive normalized numeric code from filename
      - try exact match on NORM_CODE
      - if no match, try trimming last 1–3 digits from the numeric code
    """
    price_map = {str(r["NORM_CODE"]): r for _, r in price_df.iterrows()}

    rows = []
    for fname in photo_names:
        norm = extract_photo_norm_code(fname)

        matched_row = price_map.get(norm)

        if matched_row is None and norm:
            # Try trimming last few digits (handles variant codes)
            for trim in range(1, 4):
                if len(norm) - trim < 4:
                    break
                candidate = norm[:-trim]
                matched_row = price_map.get(candidate)
                if matched_row is not None:
                    break

        if matched_row is not None:
            rows.append({
                "PHOTO": fname,
                "CODE": matched_row["CODE"],
                "DESCRIPTION": matched_row["DESCRIPTION"],
                "PRICE_INCL": matched_row["PRICE_INCL"],
                "FILENAME": fname,
            })
        else:
            rows.append({
                "PHOTO": fname,
                "CODE": norm,
                "DESCRIPTION": "",
                "PRICE_INCL": None,
                "FILENAME": fname,
            })

    return pd.DataFrame(rows)


# ---------------------------------------------------------
# Excel builder – thumbnails + A–E columns
#   Col A: Photo thumbnail
#   Col B: Code
#   Col C: Description
#   Col D: Price incl
#   Col E: Filename
# ---------------------------------------------------------
def build_excel_with_thumbnails(df: pd.DataFrame, photo_folder: str) -> BytesIO:
    wb = Workbook()
    ws = wb.active
    ws.title = "Catalogue"

    # Header row
    ws.append(["Photo", "Code", "Description", "Price incl", "Filename"])

    # Column widths
    ws.column_dimensions["A"].width = 30
    ws.column_dimensions["B"].width = 18
    ws.column_dimensions["C"].width = 40
    ws.column_dimensions["D"].width = 14
    ws.column_dimensions["E"].width = 30

    row_idx = 2
    for _, row in df.iterrows():
        code = row.get("CODE", "")
        desc = row.get("DESCRIPTION", "")
        price = row.get("PRICE_INCL", None)
        fname = row.get("FILENAME", "")
        photo_name = row.get("PHOTO", "")

        ws.cell(row=row_idx, column=2, value=str(code))
        ws.cell(row=row_idx, column=3, value=str(desc))
        ws.cell(row=row_idx, column=5, value=str(fname))

        if isinstance(price, (int, float)) and not pd.isna(price):
            ws.cell(row=row_idx, column=4, value=float(price))

        img_path = os.path.join(photo_folder, photo_name)

        if os.path.exists(img_path):
            try:
                # Build a thumbnail and embed as JPEG to avoid MPO issues
                with Image.open(img_path) as im:
                    im = im.convert("RGB")
                    im.thumbnail((150, 150))

                    img_bytes = BytesIO()
                    im.save(img_bytes, format="JPEG")
                    img_bytes.seek(0)

                    xl_img = XLImage(img_bytes)
                    xl_img.width = 80
                    xl_img.height = 80
                    ws.add_image(xl_img, f"A{row_idx}")

                ws.row_dimensions[row_idx].height = 70
            except Exception as e:
                print("Excel thumbnail error:", e)

        row_idx += 1

    out = BytesIO()
    wb.save(out)
    out.seek(0)
    return out


# ---------------------------------------------------------
# PDF builder – 3×3 grid: photo + desc + price + code number
# ---------------------------------------------------------
def generate_pdf_3x3(df: pd.DataFrame, photo_folder: str) -> bytes:
    """
    3×3 A4 layout.
    For each item:
      [Photo]
      Description
      Price: Rxx.xx
      Code: 8613900012
    """
    pdf = FPDF("P", "mm", "A4")
    pdf.set_auto_page_break(auto=True, margin=10)

    page_w, page_h = 210, 297
    margin_x, margin_y = 10, 10
    usable_w = page_w - 2 * margin_x
    usable_h = page_h - 2 * margin_y

    cols, rows = 3, 3
    cell_w = usable_w / cols
    cell_h = usable_h / rows

    img_w = 60
    img_h = 40

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
            desc = str(row.get("DESCRIPTION", "") or "")
            code = str(row.get("CODE", "") or "")
            price_val = row.get("PRICE_INCL", None)

            if isinstance(price_val, (int, float)) and not pd.isna(price_val):
                price_str = f"R{price_val:,.2f}"
            else:
                price_str = ""

            # Draw image if exists
            if os.path.exists(img_path):
                try:
                    with Image.open(img_path) as im:
                        im = im.convert("RGB")
                        # create temp JPEG
                        tmp_bytes = BytesIO()
                        im.save(tmp_bytes, format="JPEG")
                        tmp_bytes.seek(0)

                        tmp_dir = tempfile.gettempdir()
                        tmp_img_path = os.path.join(tmp_dir, f"proto_tmp_{i}_{idx}.jpg")
                        with open(tmp_img_path, "wb") as f:
                            f.write(tmp_bytes.read())

                    img_x = x + (cell_w - img_w) / 2
                    pdf.image(tmp_img_path, x=img_x, y=y, w=img_w)
                except Exception as e:
                    print("PDF image error:", e)

            # Text area under image
            text_y = y + img_h + 3

            # Description
            pdf.set_xy(x, text_y)
            pdf.set_font("Arial", size=8)
            pdf.multi_cell(cell_w, 4, desc[:160], 0, "L")

            # Price
            if price_str:
                pdf.set_xy(x, text_y + 14)
                pdf.set_font("Arial", size=8)
                pdf.cell(cell_w, 4, f"Price: {price_str}", 0, 2, "L")

            # Code (photo number) – always shown
            if code:
                pdf.set_xy(x, text_y + 20)
                pdf.set_font("Arial", "B", 8)
                pdf.cell(cell_w, 4, f"Code: {code}", 0, 2, "L")

        # Page number
        pdf.set_y(page_h - 15)
        pdf.set_font("Arial", "I", 8)
        pdf.cell(0, 10, f"Page {pdf.page_no()}", 0, 0, "C")

    out = BytesIO()
    pdf.output(out)
    return out.getvalue()


# ---------------------------------------------------------
# Streamlit UI
# ---------------------------------------------------------
st.header("1️⃣ Upload Price PDF (PRODUCT DETAILS - BY CODE.pdf)")
price_pdf = st.file_uploader("Price PDF", type=["pdf"])

st.header("2️⃣ Upload Product Photos (JPG/JPEG/PNG)")
photos = st.file_uploader(
    "Product photos (filenames must contain the code)",
    type=["jpg", "jpeg", "png"],
    accept_multiple_files=True,
)

if st.button("🔄 Generate Excel + PDF"):
    if not price_pdf or not photos:
        st.error("Please upload BOTH the price PDF and at least one photo.")
    else:
        # 1) Extract prices
        with st.spinner("Extracting prices from PDF…"):
            pdf_bytes = price_pdf.read()
            price_df = extract_prices_from_pdf(pdf_bytes)

        st.success(f"Extracted {len(price_df)} price rows from PDF.")

        # 2) Save photos & match
        with st.spinner("Saving photos and matching to price list…"):
            temp_dir = tempfile.mkdtemp(prefix="proto_photos_")
            photo_names = []
            for pf in photos:
                photo_names.append(pf.name)
                save_path = os.path.join(temp_dir, pf.name)
                with open(save_path, "wb") as f:
                    f.write(pf.getvalue())

            matched_df = match_photos_to_prices(photo_names, price_df)

        st.subheader("Matched Data (Preview)")
        st.dataframe(matched_df)

        # 3) Build Excel
        with st.spinner("Building Excel file with thumbnails…"):
            excel_file = build_excel_with_thumbnails(matched_df, temp_dir)

        # 4) Build PDF 3×3 (photo + desc + price + CODE)
        with st.spinner("Building 3×3 PDF catalogue…"):
            pdf_file = generate_pdf_3x3(matched_df, temp_dir)

        st.success("Done! Download your files:")

        st.download_button(
            "📥 Download Excel Catalogue",
            data=excel_file,
            file_name="catalogue.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        st.download_button(
            "📄 Download 3×3 PDF Catalogue",
            data=pdf_file,
            file_name="catalogue.pdf",
            mime="application/pdf",
        )
