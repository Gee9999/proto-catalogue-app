import streamlit as st
import pandas as pd
import os
import re
import pdfplumber
from PIL import Image
from io import BytesIO
from fpdf import FPDF
import tempfile

# ---------------------------------------------------------
# Streamlit app config
# ---------------------------------------------------------
st.set_page_config(page_title="Proto Catalogue – PDF Only", layout="wide")
st.title("📸 Proto Trading – PDF Catalogue Builder (3×3, Description + Price)")


# ---------------------------------------------------------
# Helpers
# ---------------------------------------------------------
def normalize_code(code: str) -> str:
    """Keep only digits from a code."""
    return re.sub(r"[^0-9]", "", str(code))


def extract_photo_norm_code(filename: str) -> str:
    """
    Extract numeric code from photo filename.
    Example:
        8613900012-20PCS.jpg   -> 8613900012
        86101000001-10mm-10pcs -> 86101000001
    """
    stem = os.path.splitext(filename)[0]
    base = stem.split("-")[0]  # everything before first '-'
    return normalize_code(base)


# ---------------------------------------------------------
# Extract prices from PDF using pdfplumber (text-only)
# ---------------------------------------------------------
@st.cache_data
def extract_prices_from_pdf(pdf_bytes: bytes) -> pd.DataFrame:
    """
    Extract CODE, DESCRIPTION, PRICE_INCL from 'PRODUCT DETAILS - BY CODE.pdf'
    Logic (same as before):
      - Each product line starts with a code: 8613900012 or 8613900012N
      - There are at least 5 decimal numbers on the line
      - The 5th decimal number is PRICE-A INCL
      - Description is the text between the code and the first number
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

                # Must start with a code
                m = code_pattern.match(line)
                if not m:
                    continue

                code_raw = m.group(1)
                norm_code = normalize_code(code_raw)

                # Avoid duplicates – keep first occurrence
                if norm_code in items:
                    continue

                # Find all decimal numbers on the line
                numbers = re.findall(r"\d+\.\d+", line)
                if len(numbers) < 5:
                    continue

                # PRICE-A INCL = 5th number (index 4)
                try:
                    price_incl = float(numbers[4])
                except ValueError:
                    continue

                # Description = text from after the CODE until the first number
                parts = line.split()
                desc_tokens = []
                for p in parts[1:]:
                    if re.match(r"\d+\.\d+", p):
                        break
                    desc_tokens.append(p)

                description = " ".join(desc_tokens)

                items[norm_code] = {
                    "CODE": code_raw,
                    "NORM_CODE": norm_code,
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
      - derive normalized numeric code
      - try exact match on NORM_CODE
      - if not found, try trimming last 1–3 digits (for A/B suffix logic)
    """
    price_map = {str(r["NORM_CODE"]): r for _, r in price_df.iterrows()}

    rows = []
    for fname in photo_names:
        norm = extract_photo_norm_code(fname)

        matched_row = price_map.get(norm)

        # Fallback: trim last 1–3 digits to handle variants / mis-scans
        if matched_row is None and norm:
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
            })
        else:
            # No match found – still include photo (price empty)
            rows.append({
                "PHOTO": fname,
                "CODE": norm,
                "DESCRIPTION": "",
                "PRICE_INCL": None,
            })

    return pd.DataFrame(rows)


# ---------------------------------------------------------
# Build 3×3 PDF catalogue (photo + description + price)
# ---------------------------------------------------------
def generate_pdf_3x3(df: pd.DataFrame, photo_folder: str) -> bytes:
    """
    3×3 layout:
      - A4 portrait
      - 3 columns x 3 rows
      - Each cell:
          [IMAGE centered]
          Description
          Price: Rxx.xx

    No barcode printed on PDF.
    """
    pdf = FPDF("P", "mm", "A4")
    pdf.set_auto_page_break(auto=True, margin=10)

    cols, rows = 3, 3
    page_w, page_h = 210, 297
    margin_x, margin_y = 10, 10
    usable_w = page_w - 2 * margin_x
    usable_h = page_h - 2 * margin_y

    cell_w = usable_w / cols
    cell_h = usable_h / rows

    # Image size inside each cell
    img_w = 60
    img_h = 45  # approximate

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
            price_val = row["PRICE_INCL"]

            if isinstance(price_val, (int, float)) and not pd.isna(price_val):
                price_str = f"R{price_val:,.2f}"
            else:
                price_str = ""

            # Draw image if exists
            if os.path.exists(img_path):
                try:
                    # Ensure valid image format by opening & re-saving as JPEG in memory
                    with Image.open(img_path) as im:
                        im = im.convert("RGB")
                        tmp_bytes = BytesIO()
                        im.save(tmp_bytes, format="JPEG")
                        tmp_bytes.seek(0)

                        # FPDF needs a temp file for in-memory image
                        tmp_dir = tempfile.gettempdir()
                        tmp_img_path = os.path.join(tmp_dir, f"__proto_tmp_{i}_{idx}.jpg")
                        with open(tmp_img_path, "wb") as f:
                            f.write(tmp_bytes.read())

                    # Center image horizontally within the cell
                    img_x = x + (cell_w - img_w) / 2
                    pdf.image(tmp_img_path, x=img_x, y=y, w=img_w)
                except Exception as e:
                    print("PDF image error:", e)

            # Text area below image
            text_y = y + img_h + 4

            # Description (wrap)
            pdf.set_xy(x, text_y)
            pdf.set_font("Arial", size=8)
            pdf.multi_cell(cell_w, 4, desc[:160], 0, "L")

            # Price line
            if price_str:
                pdf.set_xy(x, text_y + 14)
                pdf.set_font("Arial", size=9)
                pdf.cell(cell_w, 4, f"Price: {price_str}", 0, 2, "L")

        # Optional: page number at bottom center
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

if st.button("GENERATE PDF CATALOGUE"):
    if not price_pdf or not photos:
        st.error("Please upload BOTH the price PDF and at least one photo.")
    else:
        # 1) Extract prices from PDF
        with st.spinner("Extracting prices from PDF (this may take a bit for 550 pages)…"):
            pdf_bytes = price_pdf.read()
            price_df = extract_prices_from_pdf(pdf_bytes)

        st.success(f"Extracted {len(price_df)} price lines from PDF.")

        # 2) Save photos to temp folder and match
        with st.spinner("Saving photos and matching them to prices…"):
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

        # 3) Build 3×3 PDF catalogue (photo + description + price, no barcode)
        with st.spinner("Building 3×3 PDF catalogue…"):
            pdf_file = generate_pdf_3x3(matched_df, temp_dir)

        st.success("Done! Download your PDF catalogue below:")

        st.download_button(
            "📄 Download 3×3 PDF Catalogue",
            data=pdf_file,
            file_name="product_catalogue.pdf",
            mime="application/pdf",
        )
