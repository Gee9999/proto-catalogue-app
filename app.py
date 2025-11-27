import streamlit as st
import pandas as pd
import fitz  # PyMuPDF for fast PDF text extraction
import re
import io
from PIL import Image
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.utils import ImageReader
import tempfile

st.set_page_config(page_title="Proto Catalogue Builder", layout="wide")

# -------------------------
# FAST PRICE EXTRACTION
# -------------------------
@st.cache_data
def extract_prices_fast(pdf_bytes):
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    pattern = re.compile(r"^(\d{5,20})\b")  # CODE at line-start

    items = []
    for page in doc:
        text = page.get_text()
        if not text:
            continue

        for line in text.split("\n"):
            line = line.strip()
            if not line:
                continue

            # detect code
            m = pattern.match(line)
            if not m:
                continue

            code = m.group(1)

            # find all decimal numbers
            nums = re.findall(r"\d+\.\d+", line)
            if len(nums) < 5:
                continue

            price_incl = nums[4]  # Price-A INCL (5th number)

            # extract description (words before the first number)
            parts = line.split()
            desc_parts = []
            for p in parts[1:]:
                if re.match(r"\d+\.\d+", p):
                    break
                desc_parts.append(p)
            desc = " ".join(desc_parts)

            items.append([code, desc, price_incl])

    df = pd.DataFrame(items, columns=["CODE", "DESCRIPTION", "PRICE_INCL"])
    df = df.drop_duplicates(subset=["CODE"])

    return df


# -------------------------
# CLEAN CODE from filename
# -------------------------
def clean_code_from_filename(fn):
    # e.g. 86101000001-10mm-10pcs.jpg → 86101000001
    m = re.match(r"(\d+)", fn)
    return m.group(1) if m else ""


# -------------------------
# MATCH PHOTOS TO PRICES
# -------------------------
def match_photos(photo_files, price_df):
    rows = []

    for file in photo_files:
        fn = file.name
        base_code = clean_code_from_filename(fn)

        # variants allowed: -A, -B, 4A, 4B, -4 etc.
        possible = [
            base_code,
            base_code + "A",
            base_code + "B",
            base_code + "a",
            base_code + "b",
            base_code[:-1],  # sometimes last digit dropped
        ]

        found = None
        for pcode in possible:
            if pcode in price_df["CODE"].values:
                found = pcode
                break

        if found:
            row = price_df.loc[price_df["CODE"] == found].iloc[0]
            desc = row["DESCRIPTION"]
            price = row["PRICE_INCL"]
        else:
            desc = ""
            price = ""

        img = Image.open(file).convert("RGB")
        rows.append([img, base_code, desc, price, fn])

    df = pd.DataFrame(rows, columns=["IMAGE", "CODE", "DESCRIPTION", "PRICE_INCL", "FILENAME"])
    return df


# -------------------------
# BUILD PDF CATALOGUE (3×3)
# -------------------------
def build_pdf_catalog(df):
    buffer = io.BytesIO()
    c = canvas.Canvas(buffer, pagesize=A4)
    page_w, page_h = A4

    cols = 3
    rows = 3
    cell_w = page_w / cols
    cell_h = page_h / rows

    items = df.to_dict("records")

    for idx, row in enumerate(items):
        col = idx % cols
        rw = (idx // cols) % rows

        if idx % 9 == 0 and idx != 0:
            c.showPage()

        x0 = col * cell_w
        y0 = page_h - (rw + 1) * cell_h

        pil_img = row["IMAGE"]
        max_w, max_h = 150, 150
        pil_img.thumbnail((max_w, max_h))
        img_w, img_h = pil_img.size

        img_x = x0 + (cell_w - img_w) / 2
        img_y = y0 + cell_h - img_h - 20

        c.drawImage(ImageReader(pil_img), img_x, img_y, width=img_w, height=img_h, preserveAspectRatio=True)

        # TEXT
        price_str = f"R{row['PRICE_INCL']}" if row["PRICE_INCL"] else "N/A"
        desc = row["DESCRIPTION"] if row["DESCRIPTION"] else ""
        code = row["CODE"] if row["CODE"] else ""

        text_y = img_y - 16

        # Price
        c.setFont("Helvetica", 9)
        c.drawString(x0 + 10, text_y, f"Price: {price_str}")

        # Description
        c.setFont("Helvetica", 8)
        c.drawString(x0 + 10, text_y - 14, desc[:80])

        # Code LAST LINE (Option C)
        c.setFont("Helvetica", 8)
        c.drawString(x0 + 10, text_y - 28, f"Code: {code}")

    c.showPage()
    c.save()
    buffer.seek(0)
    return buffer


# -------------------------
# STREAMLIT UI
# -------------------------

st.title("🟥 Proto Trading Catalogue Builder")

st.header("1️⃣ Upload Price PDF")
price_pdf = st.file_uploader("Upload price PDF", type=["pdf"])

st.header("2️⃣ Upload Product Photos")
photos = st.file_uploader(
    "Upload photos",
    accept_multiple_files=True,
    type=["jpg", "jpeg", "png"],
)

if price_pdf and photos:
    st.success("Files uploaded! Extracting prices…")

    price_bytes = price_pdf.read()
    price_df = extract_prices_fast(price_bytes)

    st.write(f"📘 Extracted **{len(price_df)}** price records.")

    st.success("Matching photos…")
    matched_df = match_photos(photos, price_df)

    st.write(f"🖼 Matched **{matched_df['DESCRIPTION'].count()}** photos with prices")

    # Excel thumbnail creation
    thumb_rows = []
    for _, r in matched_df.iterrows():
        img = r["IMAGE"].copy()
        img.thumbnail((200, 200))
        buf = io.BytesIO()
        img.save(buf, format="JPEG")
        buf.seek(0)
        thumb_rows.append([buf.getvalue(), r["CODE"], r["DESCRIPTION"], r["PRICE_INCL"], r["FILENAME"]])

    excel_df = pd.DataFrame(thumb_rows, columns=["THUMBNAIL", "CODE", "DESCRIPTION", "PRICE_INCL", "FILENAME"])

    # Make thumbnail a PNG in Excel
    output_excel = io.BytesIO()
    with pd.ExcelWriter(output_excel, engine="openpyxl") as writer:
        excel_df.drop(columns=["THUMBNAIL"]).to_excel(writer, index=False, sheet_name="Data")

    output_excel.seek(0)

    # PDF Build
    st.success("Building PDF Catalogue…")
    pdf_buffer = build_pdf_catalog(matched_df)

    st.download_button("⬇️ Download Excel", data=output_excel, file_name="catalogue.xlsx")
    st.download_button("⬇️ Download PDF Catalogue", data=pdf_buffer, file_name="catalogue.pdf", mime="application/pdf")

