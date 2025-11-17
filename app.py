import os
import sys
import pandas as pd
import barcode
import tempfile
import base64
from barcode.writer import ImageWriter
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import inch
from PIL import Image
import streamlit as st
import streamlit.components.v1 as components
from textwrap import wrap
import hashlib

# ----------------------------
# Setup
# ----------------------------
if getattr(sys, 'frozen', False):
    base_path = sys._MEIPASS
else:
    base_path = os.path.dirname(__file__)

downloads_folder = tempfile.gettempdir()
barcode_folder = os.path.join(base_path, "barcodes")
os.makedirs(barcode_folder, exist_ok=True)

excel_file = os.path.join(base_path, "sku_list.xlsx")
if not os.path.exists(excel_file):
    st.error(f"⚠️ Excel file not found at {excel_file}")
    sys.exit(1)

try:
    df = pd.read_excel(excel_file, dtype={"SKU": str})
    df["SKU_normalized"] = df["SKU"].astype(str).str.strip()
except Exception as e:
    st.error(f"⚠️ Failed to load Excel file: {e}")
    sys.exit(1)

# ----------------------------
# HELPERS
# ----------------------------

def safe_filename(s: str) -> str:
    """Generate a consistent filename for caching."""
    h = hashlib.sha1(s.encode("utf-8")).hexdigest()[:10]
    return f"sku_{h}"

def transform_sku_for_barcode(sku: str) -> str:
    """Apply special encoding rules (Option A)."""

    if sku.startswith("999."):
        # Add exactly two 9s to the beginning
        return "99" + sku

    return sku

# ----------------------------
# BARCODE GENERATION
# ----------------------------
def generate_barcode(sku: str):
    """Generate Code128 barcode image (with no printed text)."""

    original_sku = sku.strip()
    encoded_sku = transform_sku_for_barcode(original_sku)

    file_base = safe_filename(original_sku)
    png_path = os.path.join(barcode_folder, file_base + ".png")

    # use cached file if exists
    if os.path.exists(png_path):
        return png_path, encoded_sku

    code128 = barcode.get_barcode_class("code128")

    try:
        b_obj = code128(encoded_sku, writer=ImageWriter())

        b_obj.save(
            os.path.join(barcode_folder, file_base),
            options={
                "module_width": 0.22,
                "module_height": 8,
                "write_text": False  # IMPORTANT: do NOT show encoded text
            }
        )

    except Exception as e:
        st.error(f"❌ Barcode generation failed: {e}")
        return None, encoded_sku

    return png_path, encoded_sku

# ----------------------------
# PDF LABEL GENERATOR
# ----------------------------
def create_label_pdf(sku: str, description: str):
    pdf_path = os.path.join(downloads_folder, f"{sku}.pdf")
    c = canvas.Canvas(pdf_path, pagesize=(3 * inch, 1 * inch))

    barcode_img_path, encoded_sku = generate_barcode(sku)
    if not barcode_img_path:
        return None

    try:
        img = Image.open(barcode_img_path)
        w, h = img.size
        ratio = h / w
        target_w = 1.83 * inch
        target_h = target_w * ratio

        # Draw barcode image
        c.drawImage(
            barcode_img_path,
            x=0.6 * inch,
            y=0.45 * inch,
            width=target_w,
            height=target_h,
            preserveAspectRatio=True
        )

    except Exception as e:
        st.error(f"⚠️ Failed to load barcode image: {e}")
        return None

    # --- Draw Correct Human Readable SKU ---
    c.setFont("Helvetica", 8)
    tw = c.stringWidth(sku, "Helvetica", 8)
    c.drawString((3*inch - tw) / 2, 0.35 * inch, sku)

    # --- Draw Description ---
    if not description:
        description = ""

    c.setFont("Helvetica", 8)
    wrapped = wrap(description, 34)
    y = 0.15 * inch
    for line in wrapped:
        tw = c.stringWidth(line, "Helvetica", 8)
        c.drawString((3*inch - tw) / 2, y, line)
        y -= 0.12 * inch

    c.showPage()
    c.save()
    return pdf_path

# ----------------------------
# STREAMLIT UI
# ----------------------------
st.title("📦 SKU Barcode Generator")
st.write("Enter an SKU to generate a printable label.")

sku_input = st.text_input("Enter SKU:", "").strip()

if st.button("Generate Label"):
    if not sku_input:
        st.warning("⚠️ Please enter a valid SKU.")
        st.stop()

    match = df[df["SKU_normalized"] == sku_input]
    if match.empty:
        match = df[df["SKU_normalized"].str.lower() == sku_input.lower()]

    if match.empty:
        st.warning("⚠️ SKU not found. Try again.")
        st.stop()

    description = str(match.iloc[0].get("Description", ""))

    pdf_path = create_label_pdf(sku_input, description)

    if not pdf_path or not os.path.exists(pdf_path):
        st.error("❌ Failed to generate label.")
        st.stop()

    with open(pdf_path, "rb") as f:
        pdf_bytes = f.read()

    base64_pdf = base64.b64encode(pdf_bytes).decode("utf-8")

    print_button_html = f"""
    <script>
    function openPDF() {{
        const pdfData = atob("{base64_pdf}");
        const arr = new Uint8Array(pdfData.length);
        for (let i = 0; i < pdfData.length; i++) {{
            arr[i] = pdfData.charCodeAt(i);
        }}
        const blob = new Blob([arr], {{ type: 'application/pdf' }});
        const url = URL.createObjectURL(blob);
        const win = window.open(url);

        win.onload = function() {{
            setTimeout(() => {{
                win.focus();
                win.print();
            }}, 600);
        }};
    }}
    </script>
    <button onclick="openPDF()">🖨️ Print Label</button>
    """

    st.success("✅ Label generated!")
    components.html(print_button_html, height=120)
