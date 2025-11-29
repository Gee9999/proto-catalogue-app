import streamlit as st
import pandas as pd
import os
import tempfile
from PIL import Image
from fpdf import FPDF

# ============================================================
# HELPERS
# ============================================================

def extract_base_code(filename):
    """Extract leading digits until first non-digit."""
    base = ""
    for ch in filename:
        if ch.isdigit():
            base += ch
        else:
            break
    return base


def load_price_excel(uploaded_file):
    """Reads Excel and attempts to auto-detect columns: CODE, DESCRIPTION, PRICE."""
    df = pd.read_excel(uploaded_file, dtype=str)

    # Normalise column names
    df.columns = [c.strip().upper().replace(" ", "").replace("-", "") for c in df.columns]

    # Find CODE
    code_col = next((c for c in df.columns if "CODE" in c), None)
    # Find DESCRIPTION
    desc_col = next((c for c in df.columns if "DESC" in c or "DESCRIPTION" in c), None)
    # Find PRICE-A INCL
    price_col = next((c for c in df.columns if "PRICEAINCL" in c or "PRICE_A_INCL" in c or "INCL" in c), None)

    if not code_col or not desc_col or not price_col:
        raise ValueError("Excel must contain CODE, DESCRIPTION and PRICE-A INCL columns.")

    df = df[[code_col, desc_col, price_col]].copy()
    df.columns = ["CODE", "DESCRIPTION", "PRICE"]

    df["CODE_STR"] = df["CODE"].astype(str).str.strip()
    df = df.drop_duplicates(subset="CODE_STR")

    return df


def match_photos_to_prices(photo_files, price_df):
    """Match images to closest exact code."""
    price_lookup = price_df.set_index("CODE_STR")[["CODE", "DESCRIPTION", "PRICE"]].to_dict("index")

    rows = []

    for p in photo_files:
        fname = p.name
        base = extract_base_code(fname)

        if base in price_lookup:
            price_row = price_lookup[base]
            rows.append({
                "PHOTO_FILE": fname,
                "CODE": price_row["CODE"],
                "DESCRIPTION": price_row["DESCRIPTION"],
                "PRICE": price_row["PRICE"]
            })
        else:
            rows.append({
                "PHOTO_FILE": fname,
                "CODE": base,
                "DESCRIPTION": "NOT FOUND",
                "PRICE": "N/A"
            })

    return pd.DataFrame(rows)


# ============================================================
# EXCEL BUILDER
# ============================================================

def build_excel(df, temp_dir, thumb_small=True):
    """Create Excel with thumbnails."""
    from openpyxl import Workbook
    from openpyxl.drawing.image import Image as XLImage

    wb = Workbook()
    ws = wb.active
    ws.title = "Catalogue"

    # headers
    ws.cell(1, 1, "PHOTO")
    ws.cell(1, 2, "CODE")
    ws.cell(1, 3, "DESCRIPTION")
    ws.cell(1, 4, "PRICE")
    ws.cell(1, 5, "FILENAME")

    thumb_size = (60, 60) if thumb_small else (120, 120)

    row_idx = 2
    for _, r in df.iterrows():
        photo_path = os.path.join(temp_dir, r["PHOTO_FILE"])

        # Thumbnail
        try:
            im = Image.open(photo_path)
            im.thumbnail(thumb_size)
            thumb_path = os.path.join(temp_dir, f"thumb_{r['PHOTO_FILE']}.jpg")
            im.save(thumb_path)

            xlimg = XLImage(thumb_path)
            xlimg.width = thumb_size[0]
            xlimg.height = thumb_size[1]
            ws.add_image(xlimg, f"A{row_idx}")
        except:
            pass

        ws.cell(row_idx, 2, r["CODE"])
        ws.cell(row_idx, 3, r["DESCRIPTION"])
        ws.cell(row_idx, 4, r["PRICE"])
        ws.cell(row_idx, 5, r["PHOTO_FILE"])
        row_idx += 1

    out = os.path.join(temp_dir, "output.xlsx")
    wb.save(out)
    return out


# ============================================================
# PDF BUILDER (3×3 GRID)
# ============================================================

def build_pdf(df, temp_dir, cell_w=63, cell_h=63):
    pdf = FPDF(unit="mm", format="A4")
    pdf.add_page()
    pdf.set_auto_page_break(auto=True, margin=10)

    x_start = 10
    y_start = 20
    x = x_start
    y = y_start
    count = 0

    for _, row in df.iterrows():
        photo_path = os.path.join(temp_dir, row["PHOTO_FILE"])

        # photo
        try:
            img = Image.open(photo_path)
            img.thumbnail((cell_w, cell_h))
            temp_img = os.path.join(temp_dir, f"pdf_{row['PHOTO_FILE']}.jpg")
            img.save(temp_img)
            pdf.image(temp_img, x=x, y=y, w=cell_w, h=cell_h)
        except:
            pass

        # text below image
        pdf.set_xy(x, y + cell_h + 2)
        pdf.set_font("Arial", size=8)
        pdf.multi_cell(cell_w, 4, f"{row['CODE']}", 0, "L")
        pdf.multi_cell(cell_w, 4, f"{row['DESCRIPTION']}", 0, "L")
        pdf.multi_cell(cell_w, 4, f"Price: {row['PRICE']}", 0, "L")

        # move cursor
        x += cell_w + 10
        count += 1

        # new row of 3
        if count % 3 == 0:
            x = x_start
            y += cell_h + 25

        # new page
        if y > 250:
            pdf.add_page()
            x = x_start
            y = y_start

    # FIX ENCODING ISSUE
    out = pdf.output(dest="S")
    return bytes(out) if isinstance(out, (bytes, bytearray)) else out.encode("latin1")


# ============================================================
# STREAMLIT UI
# ============================================================

def main():
    st.title("📸 Photo Catalogue Builder")

    st.write("Upload Excel with CODE, DESCRIPTION, and VAT-inclusive price.")
    price_file = st.file_uploader("Upload Price Excel", type=["xls", "xlsx"])

    st.write("Upload photos")
    photos = st.file_uploader("Upload Photos", type=["jpg", "jpeg", "png"], accept_multiple_files=True)

    if st.button("Generate Catalogue"):
        if not price_file or not photos:
            st.error("Please upload Excel + product photos")
            return

        with tempfile.TemporaryDirectory() as temp_dir:
            # Save photos temporarily
            for p in photos:
                with open(os.path.join(temp_dir, p.name), "wb") as f:
                    f.write(p.read())

            # Load prices
            price_df = load_price_excel(price_file)

            # Match photos
            matched_df = match_photos_to_prices(photos, price_df)

            # Thumbnail size choice
            thumb_small = len(photos) > 200  # Auto logic

            # Build Excel
            excel_path = build_excel(matched_df, temp_dir, thumb_small=thumb_small)
            with open(excel_path, "rb") as f:
                st.download_button("⬇️ Download Excel", f, "catalogue.xlsx")

            # Build PDF
            pdf_bytes = build_pdf(matched_df, temp_dir)
            st.download_button("⬇️ Download PDF", pdf_bytes, "catalogue.pdf", mime="application/pdf")

            st.success("Catalogue generated successfully!")


if __name__ == "__main__":
    main()
