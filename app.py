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
# STEP 1 — Extract prices from PDF using regex
# -----------------------------------------
def extract_prices_from_pdf(pdf_file):
    items = []

    code_pattern = re.compile(r"^(\d+[A-Za-z]?)\b")

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

                # find all decimal numbers
                numbers = [p for p in parts if re.match(r"^\d+\.\d+$", p)]
                if len(numbers) < 5:
                    continue

                # PRICE-A INCL is 5th number
                price_incl = float(numbers[4])

                # DESCRIPTION: tokens before price numbers
                desc_tokens = []
                for p in parts[1:]:
                    if re.match(r"^\d+\.\d+$", p):
                        break
                    desc_tokens.append(p)

                description = " ".join(desc_tokens)

                items.append([code, description, price_incl])

    return pd.DataFrame(items, columns=["CODE", "DESCRIPTION", "PRICE_INCL"])


# -----------------------------------------
# STEP 2 — Match photos to codes
# -----------------------------------------
def normalize_code(code):
    return re.sub(r"[^0-9]", "", str(code))


def match_photos_to_prices(price_df, photos):
    rows = []

    for f in photos:
        filename = f.name
        base_part = filename.split("-")[0]
        base_code = normalize_code(base_part)

        matches = price_df[
            price_df["CODE"].astype(str).apply(normalize_code) == base_code
        ]

        if not matches.empty:
            desc = matches.iloc[0]["DESCRIPTION"]
            price = matches.iloc[0]["PRICE_INCL"]
        else:
            desc = "NO MATCH"
            price = ""

        rows.append({
            "FILENAME": filename,
            "CODE": base_code,
            "DESCRIPTION": desc,
            "PRICE_INCL": price
        })

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
                pil_img.thumbnail((180, 180))  # double size
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
# STEP 4 — 3x3 PDF Catalogue (NO BARCODE)
# -----------------------------------------
def build_pdf_catalog(df, photos):
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
        code = row["CODE"]
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
    type=["jpg", "jpeg", "png"]
)

if st.button("PROCESS"):
    if not price_pdf or not photos:
        st.error("Please upload both the price PDF and photos.")
    else:
        with st.spinner("Extracting prices..."):
            price_df = extract_prices_from_pdf(price_pdf)

        with st.spinner("Matching photos..."):
            matched_df = match_photos_to_prices(price_df, photos)

        # 🩹 FIX STREAMLIT ARROW ERROR
        matched_df["PRICE_INCL"] = pd.to_numeric(
            matched_df["PRICE_INCL"], errors="coerce"
        )

        with st.spinner("Building Excel..."):
            excel_file = build_excel_with_thumbnails(matched_df, photos)

        with st.spinner("Building PDF..."):
            pdf_file = build_pdf_catalog(matched_df, photos)

        st.success("Done! Your files are ready.")

        st.download_button(
            "⬇️ Download Excel",
            data=excel_file,
            file_name="product_catalogue.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        st.download_button(
            "📄 Download 3x3 PDF Catalogue",
            data=pdf_file,
            file_name="product_catalogue.pdf",
            mime="application/pdf"
        )

        st.subheader("Preview of Matched Data")
        st.dataframe(matched_df)
