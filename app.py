import streamlit as st
import pandas as pd
import os
import re
import fitz            # PyMuPDF
import pytesseract     # OCR
from PIL import Image
from io import BytesIO
from fpdf import FPDF

# Ensure Tesseract path exists on Windows
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

st.title("📸 Product Photo + Price Catalogue Generator (OCR Enabled)")

# -----------------------------
# EXTRACT TEXT (FAST MODE)
# -----------------------------
def extract_text_fast(doc):
    text = ""
    for page in doc:
        try:
            text += page.get_text()
        except:
            pass
    return text

# -----------------------------
# EXTRACT TEXT WITH OCR (SLOW MODE)
# -----------------------------
def extract_text_ocr(doc):
    text = ""
    for page_no, page in enumerate(doc, start=1):
        pix = page.get_pixmap(dpi=200)
        img = Image.open(BytesIO(pix.tobytes("png")))
        page_text = pytesseract.image_to_string(img)
        text += page_text
    return text

# -----------------------------
# PARSE PRICE LINES
# -----------------------------
price_regex = re.compile(
    r"(?P<code>\d{6,12}[A-Za-z]?)\s+(?P<desc>.*?)\s+(?P<price>\d+\.\d{2})"
)

def extract_prices(text):
    rows = []
    for line in text.split("\n"):
        line = line.strip()
        m = price_regex.search(line)
        if m:
            rows.append([
                m.group("code"),
                m.group("desc"),
                m.group("price")
            ])
    return rows


# -----------------------------
# MATCH PHOTO → PRICE
# -----------------------------
def normalize_code(x):
    return re.sub(r"[^0-9]", "", x)

def match_photos(photo_files, price_df):

    price_df["NORM"] = price_df["CODE"].apply(normalize_code)
    rows = []

    for fname in photo_files:
        name_no_ext = os.path.splitext(fname)[0]
        clean = normalize_code(name_no_ext)

        match = price_df[price_df["NORM"] == clean]

        if not match.empty:
            desc = match.iloc[0]["DESCRIPTION"]
            price = match.iloc[0]["PRICE"]
        else:
            desc = ""
            price = ""

        rows.append([fname, clean, desc, price, fname])

    df = pd.DataFrame(rows, columns=["PHOTO", "CODE", "DESCRIPTION", "PRICE", "FILENAME"])
    return df


# -----------------------------
# PHOTO → PDF (3×3 GRID)
# -----------------------------
def generate_pdf(df, photo_folder):
    pdf = FPDF("P", "mm", "A4")
    pdf.set_auto_page_break(auto=True, margin=10)

    # grid coordinates
    cols = 3
    rows = 3
    img_w = 60
    img_h = 60
    padding_x = 10
    padding_y = 10
    text_h = 5

    for i in range(0, len(df), 9):
        pdf.add_page()
        chunk = df.iloc[i:i + 9]

        for idx, row in chunk.iterrows():
            grid_i = idx % 9
            r = grid_i // cols
            c = grid_i % cols

            x = padding_x + c * (img_w + 10)
            y = padding_y + r * (img_h + 20)

            img_path = os.path.join(photo_folder, row["PHOTO"])
            try:
                pdf.image(img_path, x=x, y=y, w=img_w, h=img_h)
            except:
                pass

            pdf.set_xy(x, y + img_h + 2)
            pdf.set_font("Arial", size=10)
            pdf.multi_cell(img_w, text_h, f"{row['DESCRIPTION']}", 0, "L")

            # BARCODE number under description
            pdf.set_xy(x, y + img_h + 12)
            pdf.set_font("Arial", size=9)
            pdf.multi_cell(img_w, text_h, f"Code: {row['CODE']}", 0, "L")

            # PRICE
            pdf.set_xy(x, y + img_h + 20)
            pdf.set_font("Arial", size=10)
            pdf.multi_cell(img_w, text_h, f"Price: R{row['PRICE']}", 0, "L")

    output = BytesIO()
    pdf.output(output)
    return output.getvalue()


# -----------------------------
# STREAMLIT UI
# -----------------------------
st.header("1️⃣ Upload Price PDF")
pdf_file = st.file_uploader("Choose PDF", type=["pdf"])

st.header("2️⃣ Upload Photos Folder")
photo_files = st.file_uploader("Upload Photos (multiple allowed)", type=["jpg", "jpeg", "png"], accept_multiple_files=True)

if pdf_file and photo_files:
    st.success("Processing… please wait.")

    # Load PDF
    doc = fitz.open(stream=pdf_file.read(), filetype="pdf")

    # FAST extraction
    text = extract_text_fast(doc)

    # If text is too small, switch to OCR
    if len(text.strip()) < 500:
        st.warning("Normal extraction failed → using OCR mode (slower but accurate)…")
        doc = fitz.open(stream=pdf_file.read(), filetype="pdf")
        text = extract_text_ocr(doc)

    rows = extract_prices(text)
    price_df = pd.DataFrame(rows, columns=["CODE", "DESCRIPTION", "PRICE"])

    # MATCHING
    photo_names = [p.name for p in photo_files]
    temp_folder = "uploaded_photos"
    os.makedirs(temp_folder, exist_ok=True)

    for p in photo_files:
        with open(os.path.join(temp_folder, p.name), "wb") as f:
            f.write(p.read())

    result_df = match_photos(photo_names, price_df)

    st.header("3️⃣ Download Excel")
    st.dataframe(result_df)

    excel_bytes = result_df.to_excel(index=False)
    st.download_button("⬇ Download Excel", excel_bytes, "catalogue.xlsx")

    st.header("4️⃣ Generate PDF (3×3 Grid)")
    pdf_bytes = generate_pdf(result_df, temp_folder)
    st.download_button("⬇ Download PDF", pdf_bytes, "catalogue.pdf", mime="application/pdf")

