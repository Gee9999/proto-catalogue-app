import os
import re
import tempfile
from io import BytesIO

import pandas as pd
import streamlit as st
from fpdf import FPDF
from PIL import Image


# ---------- Helpers ----------

def normalize_header(col) -> str:
    """
    Make Excel column headers comparable even if they have spaces, hyphens, dots, or line breaks.
    Example:
      "PRICE-A\nINCL." -> "priceaincl"
      " Price A INCL " -> "priceaincl"
    """
    return (
        str(col)
        .strip()
        .replace("-", "")
        .replace(".", "")
        .replace(" ", "")
        .replace("\n", "")
        .lower()
    )


def normalize_code(val) -> str | None:
    """
    Turn any CODE (numeric / string / scientific notation) into a pure digit string.
    Example:
      8.6104E+09  -> "8610400000..."
      "8613900001" -> "8613900001"
      " 8613900001 " -> "8613900001"
    """
    if pd.isna(val):
        return None
    s = str(val)
    digits = "".join(ch for ch in s if ch.isdigit())
    return digits or None


def extract_code_from_filename(filename: str) -> str | None:
    """
    Your photos look like:
      8613900001-25PCS.JPG
      8613900017-15mm-20pcs.jpg
      8610400003-50pcs.JPG
    We take all leading digits from the start of the filename.
    """
    basename = os.path.basename(filename)
    m = re.match(r"^(\d+)", basename)
    if not m:
        return None
    return m.group(1)


def longest_common_prefix_len(a: str, b: str) -> int:
    """Return length of longest common prefix between two strings."""
    n = min(len(a), len(b))
    for i in range(n):
        if a[i] != b[i]:
            return i
    return n


def choose_best_code(photo_code: str, codes: pd.Series) -> str | None:
    """
    Try:
      1) Exact match on normalized code.
      2) If no exact match, pick the code with the longest common prefix.
         (Your 'base number / closest match' idea.)
    """
    if not photo_code:
        return None

    codes = codes.dropna().astype(str)

    # 1) Exact match
    exact = codes[codes == photo_code]
    if not exact.empty:
        return exact.iloc[0]

    # 2) Longest common prefix
    best_code = None
    best_score = 0
    for c in codes.unique():
        score = longest_common_prefix_len(photo_code, c)
        if score > best_score:
            best_score = score
            best_code = c

    if best_score == 0:
        return None

    return best_code


def detect_columns(df: pd.DataFrame) -> tuple[str, str, str]:
    """
    Auto-detect which columns are CODE, DESCRIPTION, and PRICE-A INCL (or equivalent).
    Works even if headers are a bit messy.
    """
    norm_cols = [normalize_header(c) for c in df.columns]

    code_candidates = {"code", "itemcode", "barcode", "barcodenumber"}
    desc_candidates = {"description", "desc"}
    price_candidates = {"priceaincl", "pricea", "priceincl", "sellingpriceincl", "sellpriceincl"}

    code_col = None
    desc_col = None
    price_col = None

    for original, norm in zip(df.columns, norm_cols):
        if code_col is None and norm in code_candidates:
            code_col = original
        if desc_col is None and norm in desc_candidates:
            desc_col = original
        if price_col is None and norm in price_candidates:
            price_col = original

    # If any still None, fall back to fuzzy contains
    if code_col is None:
        for original, norm in zip(df.columns, norm_cols):
            if "code" in norm or "barcode" in norm:
                code_col = original
                break

    if desc_col is None:
        for original, norm in zip(df.columns, norm_cols):
            if "description" in norm or norm.startswith("desc"):
                desc_col = original
                break

    if price_col is None:
        for original, norm in zip(df.columns, norm_cols):
            if "price" in norm and ("incl" in norm or "a" in norm):
                price_col = original
                break

    if not code_col or not desc_col or not price_col:
        raise ValueError(
            "Could not automatically detect CODE / DESCRIPTION / PRICE columns.\n"
            "Make sure your file has logical headers like CODE, DESCRIPTION, PRICE-AINCL."
        )

    return code_col, desc_col, price_col


def format_price(val) -> str:
    try:
        return f"{float(val):.2f}"
    except Exception:
        return str(val)


def build_pdf(records: list[dict]) -> bytes:
    """
    Build a PDF catalogue:
      - 3 items per row
      - image on top
      - below: code, description, price
    """
    if not records:
        return b""

    pdf = FPDF("P", "mm", "A4")
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()

    # Title
    pdf.set_font("Arial", "B", 14)
    pdf.cell(0, 10, "Photo Catalogue", ln=1, align="C")
    pdf.ln(3)

    margin = 10
    columns = 3
    page_width = pdf.w - 2 * margin
    cell_w = page_width / columns
    img_h = 40  # image height
    text_h = 12  # space for code/desc/price
    row_h = img_h + text_h + 5

    x_start = margin
    y = pdf.get_y()

    for idx, rec in enumerate(records):
        col = idx % columns
        if col == 0 and idx != 0:
            # new row
            y += row_h
            if y + row_h > (pdf.h - pdf.b_margin):
                pdf.add_page()
                y = margin

        x = x_start + col * cell_w

        # Draw image
        pdf.set_xy(x, y)
        try:
            pdf.image(rec["tmp_path"], x=x + 2, y=y, w=cell_w - 4, h=img_h)
        except Exception:
            # if image fails, just ignore but still show text
            pass

        # Text below
        pdf.set_xy(x, y + img_h + 1)
        pdf.set_font("Arial", size=8)

        code_line = f"Code: {rec.get('code_display', '')}"
        desc_line = rec.get("description", "") or ""
        price_line = f"Price A incl: {rec.get('price', '')}"

        pdf.cell(cell_w, 4, code_line[:60], ln=1)
        pdf.set_x(x)
        pdf.cell(cell_w, 4, desc_line[:60], ln=1)
        pdf.set_x(x)
        pdf.cell(cell_w, 4, price_line[:60], ln=1)

    # Return PDF as bytes
    return pdf.output(dest="S").encode("latin1")


def build_excel(records: list[dict]) -> bytes:
    """
    Create a simple Excel with:
      PHOTO, CODE, DESCRIPTION, PRICE_A_INCL, MATCH_STATUS
    (No thumbnails – avoids the .mpo / mimetype issues.)
    """
    if not records:
        df = pd.DataFrame(
            columns=["PHOTO", "CODE", "DESCRIPTION", "PRICE_A_INCL", "MATCH_STATUS"]
        )
    else:
        rows = []
        for rec in records:
            rows.append(
                {
                    "PHOTO": rec.get("photo_name", ""),
                    "CODE": rec.get("code_display", ""),
                    "DESCRIPTION": rec.get("description", ""),
                    "PRICE_A_INCL": rec.get("price", ""),
                    "MATCH_STATUS": rec.get("match_status", ""),
                }
            )
        df = pd.DataFrame(rows)

    out = BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Catalogue")
    out.seek(0)
    return out.getvalue()


# ---------- Streamlit App ----------

def main():
    st.set_page_config(page_title="Photo Catalogue Builder", layout="centered")
    st.title("📸 Photo Catalogue Builder")

    st.markdown(
        """
Upload:

1. **Price Excel** (e.g. `PRODUCT DETAILS - BY CODE.xlsx`)  
2. **Product photos** (`.jpg` / `.jpeg` / `.png`)

The app will:
- Extract the numeric code from each photo filename (e.g. `8610400003-50pcs.JPG` → `8610400003`)
- Match that to your Excel by **CODE**
- Pull **DESCRIPTION** and **PRICE A incl.**
- Generate:
  - A **PDF catalogue** with photos + code + description + price
  - An **Excel file** with the matched data
        """
    )

    st.info(
        "Excel should logically contain columns for **CODE**, **DESCRIPTION**, and **PRICE-A INCL**.\n"
        "Header text can be messy (hyphens, dots, line breaks) – I'll normalize it."
    )

    price_file = st.file_uploader(
        "Upload Excel price file",
        type=["xlsx", "xls"],
        help="Example: PRODUCT DETAILS - BY CODE.xlsx",
    )

    photo_files = st.file_uploader(
        "Upload product photos",
        type=["jpg", "jpeg", "png"],
        accept_multiple_files=True,
        help="Filenames should start with the numeric code, e.g. 8613900001-25PCS.JPG",
    )

    if not price_file or not photo_files:
        st.stop()

    if st.button("🔄 Build PDF + Excel catalogue"):
        with st.spinner("Processing..."):
            try:
                # --- 1. Read Excel ---
                df = pd.read_excel(price_file)

                code_col, desc_col, price_col = detect_columns(df)

                # Keep only the 3 key columns, standardised
                df = df[[code_col, desc_col, price_col]].copy()
                df.columns = ["CODE", "DESCRIPTION", "PRICE_RAW"]

                # Normalized code for joining
                df["CODE_KEY"] = df["CODE"].apply(normalize_code)
                df = df[df["CODE_KEY"].notna()]

                # --- 2. Process photos & match ---
                records = []

                with tempfile.TemporaryDirectory() as tmpdir:
                    all_codes = df["CODE_KEY"]

                    for uploaded in photo_files:
                        photo_name = uploaded.name
                        image_bytes = uploaded.getvalue()

                        # save to temp for PDF
                        tmp_path = os.path.join(tmpdir, photo_name)
                        with open(tmp_path, "wb") as f:
                            f.write(image_bytes)

                        photo_code = extract_code_from_filename(photo_name)
                        best_code = choose_best_code(photo_code, all_codes) if photo_code else None

                        if best_code is None:
                            # No match found
                            rec = {
                                "photo_name": photo_name,
                                "tmp_path": tmp_path,
                                "code_display": photo_code or "",
                                "description": "",
                                "price": "",
                                "match_status": "NO MATCH",
                            }
                        else:
                            row = df[df["CODE_KEY"] == best_code].iloc[0]
                            description = row.get("DESCRIPTION", "")
                            price = format_price(row.get("PRICE_RAW", ""))

                            # Decide if exact or fuzzy
                            if photo_code == best_code:
                                status = "EXACT"
                            else:
                                status = "BEST PREFIX"

                            rec = {
                                "photo_name": photo_name,
                                "tmp_path": tmp_path,
                                "code_display": str(row.get("CODE", best_code)),
                                "description": str(description),
                                "price": price,
                                "match_status": status,
                            }

                        records.append(rec)

                    # --- 3. Build PDF & Excel ---
                    pdf_bytes = build_pdf(records)
                    excel_bytes = build_excel(records)

                # --- 4. Show preview table in Streamlit ---
                preview_rows = [
                    {
                        "PHOTO": r["photo_name"],
                        "CODE": r["code_display"],
                        "DESCRIPTION": r["description"],
                        "PRICE_A_INCL": r["price"],
                        "MATCH_STATUS": r["match_status"],
                    }
                    for r in records
                ]
                preview_df = pd.DataFrame(preview_rows)
                st.subheader("Preview of matched items")
                st.dataframe(preview_df, use_container_width=True)

                # --- 5. Download buttons ---
                st.subheader("Downloads")

                st.download_button(
                    "⬇️ Download PDF catalogue",
                    data=pdf_bytes,
                    file_name="photo_catalogue.pdf",
                    mime="application/pdf",
                )

                st.download_button(
                    "⬇️ Download Excel (CODE + DESC + PRICE)",
                    data=excel_bytes,
                    file_name="photo_catalogue.xlsx",
                    mime=(
                        "application/vnd.openxmlformats-officedocument."
                        "spreadsheetml.sheet"
                    ),
                )

            except Exception as e:
                st.error(f"Something went wrong: {e}")


if __name__ == "__main__":
    main()
