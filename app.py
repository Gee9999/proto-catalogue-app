import streamlit as st
import pdfplumber
import re
import pandas as pd
from io import BytesIO
from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage
from PIL import Image as PILImage

from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader


st.title("📸 Product Catalogue Builder (Proto Trading)")


# -----------------------------------------
# Helpers
# -----------------------------------------
def normalize_code(code: str) -> str:
    """
    Keep only digits so that:
    86101000001, 86101000001A, 86101000001-10mm-10pcs
    all normalise to: 86101000001
    """
    return re.sub(r"[^0-9]", "", str(code))


def get_wanted_codes_from_photos(photos):
    """
    From photo filenames, build a set of base codes we care about.
    e.g. '86101000001-10mm-10pcs.jpg' -> '86101000001'
    """
    wanted = set()
    for f in photos:
        filename = f.name
        base_part = filename.split("-")[0]  # part before first dash
        norm = normalize_code(base_part)
        if norm:
            wanted.add(norm)
    return wanted


# -----------------------------------------
# STEP 1 — Extract ONLY the needed prices from BIG PDF
# -----------------------------------------
def extract_prices_for_codes(pdf_file, wanted_codes):
    """
    Fast mode:
    - Scan all pages
    - For each line starting with a code
    - If that code matches one of our wanted_codes (normalised),
      extract DESCRIPTION and PRICE-A INCL (5th decimal number).
    - Return a dict: norm_code -> {CODE, DESCRIPTION, PRICE_INCL}
    """
    price_info = {}

    # Code at start of line
    code_pattern = re.compile(r"^(\d+[A-Za-z]?)\b")

    with pdfplumber.open(pdf_file) as pdf:
        total_pages = len(pdf.pages)
        for page_index, page in enumerate(pdf.pages):
            text = page.extract_text()
            if not text:
                continue

            for line in text.split("\n"):
                line = line.strip()
                if not line:
                    continue

                m = code_pattern.match(line)
                if not m:
                    continue

                orig_code = m.group(1)
                norm_code = normalize_code(orig_code)

                # Skip lines for codes we don't have photos for
                if norm_code not in wanted_codes:
                    continue

                parts = line.split()

                # All decimal numbers on this row
                numbers = [p for p in parts if re.match(r"^\d+\.\d+$", p)]
                # We expect at least 5 decimals; the 5th one is PRICE-A INCL
                if len(numbers) < 5:
                    continue

                price_incl = float(numbers[4])

                # DESCRIPTION = tokens after code until first decimal number
                desc_tokens = []
                for p in parts[1:]:
                    if re.match(r"^\d+\.\d+$", p):
                        break
                    desc_tokens.append(p)
                description = " ".join(desc_tokens)

                # Only keep the first match we see for that code
                if norm_code not in price_info:
                    price_info[norm_code] = {
                        "CODE": orig_code,
                        "DESCRIPTION": description,
                        "PRICE_INCL": price_incl,
                    }

    return price_info


# -----------------------------------------
# STEP 2 — Build matched DataFrame (one row per photo)
# -----------------------------------------
def build_matched_df_from_price_info(photos, price_info):
    """
    For each photo:
    - Work out its base code (from filename)
    - Look up that code in price_info
    - Build a row with: FILENAME, CODE, DESCRIPTION, PRICE_INCL
    """
    rows = []

    for f in photos:
        filename = f.name
        base_part = filename.split("-")[0]
        norm_code = normalize_code(base_part)

        info = price_info.get(norm_code)

        if info:
            code_val = info["CODE"]
            desc = info["DESCRIPTION"]
            price = info["PRICE_INCL"]
        else:
            # No match found in PDF
            code_val = norm_code
            desc = "NO MATCH"
            price = ""

        rows.append(
            {
                "FILENAME": filename,
                "CODE": code_val,
                "DESCRIPTION": desc,
                "PRICE_INCL": price,
            }
        )

    return pd.DataFrame(rows)


# -----------------------------------------
# STEP 3 — Excel Output with Double Thumbnails
# -----------------------------------------
def build_excel_with_thumbnails(df, photos):
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

    file_dict = {f.name: f for f in photos}

    row_idx = 2
    for _, row in df.iterrows():
        filename = row["FILENAME"]
        code = row["CODE"]
        desc = row["DESCRIPTION"]
        price = row["PRICE_INCL"]

        ws.cell(row=row_idx, column=2, value=code)
        ws.cell(row=row_idx, column=3, value=desc)
        ws.cell(row=row_idx, column=5, value=filename)

        if price != "" and not pd.isna(price):
            ws.cell(row=row_idx, column=4, value=float(price))

        f = file_dict.get(filename)
        if f:
            try:
                pil_img = PILImage.open(BytesIO(f.getvalue()))
                # Double-size thumbnails
                pil_img.thumbnail((180, 180))
                img_buf = BytesIO()
                pil_img.save(img_buf, format="PNG")
                img_buf.seek(0)

                xl_img = XLImage(img_buf)
                ws.add_image(xl_img, f"A{row_idx}")
                ws.row_dimensions[row_idx].height = 140
            except Exception as e:
                print("Image load error:", e)

        row_idx += 1

    out = BytesIO()
    wb.save(out)
    out.seek(0)
    return out


# -----------------------------------------
# STEP 4 — 3x3 PDF Catalogue (no barcode)
# -----------------------------------------
def build_pdf_catalog(df, photos):
    """
    3x3 grid per page:
    - Big image
    - Price
    - Description
    - Page number at bottom
    """
    file_dict = {f.name: f for f in photos}

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

    for idx, (_, row) in enumerate(df.iterrows()):
        pos = idx % items_per_page

        if pos == 0 and idx != 0:
            c.setFont("Helvetica", 9)
            c.drawCentredString(page_w / 2, margin_bottom / 2,
                                f"Page {c.getPageNumber()}")
            c.showPage()

        col = pos % cols
        row_idx = pos // cols

        x0 = margin_left + col * cell_w
        y0 = margin_bottom + (rows - 1 - row_idx) * cell_h

        filename = row["FILENAME"]
        desc = str(row["DESCRIPTION"]) if not pd.isna(row["DESCRIPTION"]) else ""
        price = row["PRICE_INCL"]

        if isinstance(price, (int, float)) and not pd.isna(price):
            price_str = f"R{price:,.2f}"
        else:
            price_str = ""

        f = file_dict.get(filename)
        img_height_used = 0

        if f:
            try:
                pil_img = PILImage.open(BytesIO(f.getvalue()))
                pil_img.thumbnail((img_max_w, img_max_h))
                img_buf = BytesIO()
                pil_img.save(img_buf, format="PNG")
                img_buf.seek(0)

                iw, ih = pil_img.size
                img_reader = ImageReader(img_buf)
                img_x = x0 + (cell_w - iw) / 2
                img_y = y0 + cell_h - ih - 10

                c.drawImage(img_reader, img_x, img_y, width=iw, height=ih)
                img_height_used = ih
            except Exception as e:
                print("Image error:", e)

        text_y = y0 + cell_h - img_height_used - 20

        c.setFont("Helvetica", 9)
        c.drawString(x0 + 10, text_y, f"Price: {price_str}")

        c.setFont("Helvetica", 8)
        desc_line = desc[:80]
        c.drawString(x0 + 10, text_y - 14, desc_line)

    c.setFont("Helvetica", 9)
    c.drawCentredString(page_w / 2, margin_bottom / 2,
                        f"Page {c.getPageNumber()}")
    c.save()

    buf.seek(0)
    return buf


# -----------------------------------------
# STREAMLIT UI
# -----------------------------------------
st.header("1️⃣ Upload Price PDF")
price_pdf = st.file_uploader("Upload price PDF", type=["pdf"])

st.header("2️⃣ Upload Product Photos")
photos = st.file_uploader(
    "Upload photos",
    accept_multiple_files=True,
    type=["jpg", "jpeg", "png"],
)

if st.button("PROCESS"):
    if not price_pdf or not photos:
        st.error("Please upload both the price PDF and photos.")
    else:
        with st.spinner("Reading photo codes..."):
            wanted_codes = get_wanted_codes_from_photos(photos)

        with st.spinner("Extracting prices from big PDF (fast mode)..."):
            price_info = extract_prices_for_codes(price_pdf, wanted_codes)

        with st.spinner("Matching photos to prices..."):
            matched_df = build_matched_df_from_price_info(photos, price_info)

        # Make Streamlit happy with Arrow types
        matched_df["PRICE_INCL"] = pd.to_numeric(
            matched_df["PRICE_INCL"], errors="coerce"
        )

        with st.spinner("Building Excel with thumbnails..."):
            excel_file = build_excel_with_thumbnails(matched_df, photos)

        with st.spinner("Building 3x3 PDF catalogue..."):
            pdf_file = build_pdf_catalog(matched_df, photos)

        st.success("Done! Your files are ready.")

        st.download_button(
            "⬇️ Download Excel",
            data=excel_file,
            file_name="product_catalogue.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        st.download_button(
            "📄 Download 3x3 PDF Catalogue",
            data=pdf_file,
            file_name="product_catalogue.pdf",
            mime="application/pdf",
        )

        st.subheader("Preview of matched data")
        st.dataframe(matched_df)
