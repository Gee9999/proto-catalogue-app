import streamlit as st
import pandas as pd
import os
import re
import fitz              # PyMuPDF (fast PDF)
import pytesseract       # OCR fallback
from PIL import Image
from io import BytesIO
from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage
from fpdf import FPDF
import tempfile

# IMPORTANT: Point to your Tesseract installation
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

# Streamlit page
st.set_page_config(page_title="Proto Catalogue (OCR)", layout="wide")
st.title("📸 Proto Trading – OCR Catalogue Builder (3×3 / 4×4)")

# ---------------------------------------------------------
# Helper functions
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
# Extract prices from PDF (FAST + OCR fallback)
# ---------------------------------------------------------
@st.cache_data
def extract_prices_from_pdf(pdf_bytes: bytes) -> pd.DataFrame:
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    code_pattern = re.compile(r"^(\d+[A-Za-z]?)\b")

    items = {}

    for page in doc:
        # first try fast text extraction
        text = page.get_text("text")

        # if no text → use OCR
        if not text or len(text.strip()) < 10:
            pix = page.get_pixmap(dpi=200)
            img = Image.open(BytesIO(pix.tobytes("png")))
            text = pytesseract.image_to_string(img)

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

            if norm in items:
                continue

            # find decimals
            numbers = re.findall(r"\d+\.\d+", line)
            if len(numbers) < 5:
                continue

            price_incl = float(numbers[4])

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

    return pd.DataFrame(items.values())


# ---------------------------------------------------------
# Match photos to prices
# ---------------------------------------------------------
def match_photos_to_prices(photo_names, price_df):
    price_map = {str(r["NORM_CODE"]): r for _, r in price_df.iterrows()}

    rows = []
    for fname in photo_names:
        norm = extract_photo_norm_code(fname)

        row = price_map.get(norm)

        # fallback: trim last digits
        if row is None:
            for trim in range(1, 4):
                if len(norm) - trim < 4:
                    break
                candidate = norm[:-trim]
                row = price_map.get(candidate)
                if row:
                    break

        if row:
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
# Build Excel with thumbnail images
# ---------------------------------------------------------
def build_excel_with_thumbnails(df, photo_folder):
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

        if row["PRICE_INCL"] is not None:
            ws.cell(row=row_idx, column=4, value=float(row["PRICE_INCL"]))

        try:
            img_path = os.path.join(photo_folder, row["PHOTO"])
            xl_img = XLImage(img_path)
            xl_img.width = 130
            xl_img.height = 130
            ws.add_image(xl_img, f"A{row_idx}")
            ws.row_dimensions[row_idx].height = 110
        except:
            pass

        row_idx += 1

    out = BytesIO()
    wb.save(out)
    out.seek(0)
    return out


# ---------------------------------------------------------
# Build PDF catalogue (3×3 or 4×4)
# ---------------------------------------------------------
def generate_pdf_layout(df, photo_folder, layout):
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
    img_h = 45

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

            desc = row["DESCRIPTION"]
            code = row["CODE"]
            price = row["PRICE_INCL"]

            price_str = f"R{price:.2f}" if price else "R0.00"

            try:
                pdf.image(img_path, x=x + (cell_w - img_w)/2, y=y, w=img_w)
            except:
                pass

            text_y = y + img_h + 4

            pdf.set_xy(x, text_y)
            pdf.set_font("Arial", size=8)
            pdf.multi_cell(cell_w, 4, desc, 0, "L")

            pdf.set_xy(x, text_y + 16)
            pdf.set_font("Arial", size=9)
            pdf.cell(cell_w, 4, f"Price: {price_str}", 0, 2)

            pdf.set_xy(x, text_y + 23)
            pdf.cell(cell_w, 4, f"Code: {code}", 0, 2)

    out = BytesIO()
    pdf.output(out)
    return out.getvalue()


# ---------------------------------------------------------
# Streamlit UI
# ---------------------------------------------------------
st.header("1️⃣ Upload Price PDF")
pdf_file = st.file_uploader("Upload PDF", type=["pdf"])

st.header("2️⃣ Upload Photos")
photos = st.file_uploader("Upload product photos", type=["jpg","jpeg","png"], accept_multiple_files=True)

layout_option = st.radio("Choose PDF Layout:", ["3×3", "4×4"])

if st.button("PROCESS"):
    if not pdf_file or not photos:
        st.error("Upload a price PDF and photos first.")
    else:
        with st.spinner("Extracting PDF pricing… (OCR fallback applied)"):
            price_df = extract_prices_from_pdf(pdf_file.read())

        with st.spinner("Saving photos and matching…"):
            temp_dir = tempfile.mkdtemp()
            names = []
            for p in photos:
                names.append(p.name)
                with open(os.path.join(temp_dir, p.name), "wb") as f:
                    f.write(p.getvalue())

            result_df = match_photos_to_prices(names, price_df)

        st.subheader("Matched Data")
        st.dataframe(result_df)

        with st.spinner("Building Excel…"):
            excel_bytes = build_excel_with_thumbnails(result_df, temp_dir)

        with st.spinner("Building PDF…"):
            key = "3x3" if "3×3" in layout_option else "4x4"
            pdf_bytes = generate_pdf_layout(result_df, temp_dir, key)

        st.success("DONE ✔")

        st.download_button("⬇ Download Excel", data=excel_bytes,
            file_name="catalogue.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        st.download_button("⬇ Download PDF Catalogue", data=pdf_bytes,
            file_name="catalogue.pdf", mime="application/pdf")
