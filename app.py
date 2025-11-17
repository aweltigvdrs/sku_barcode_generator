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

# --------------------------------------------
# Paths & Setup
# --------------------------------------------
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

# Load Excel with SKU as string
try:
    df = pd.read_excel(excel_file, dtype={"SKU": str})
except Exception as e:
    st.error(f"⚠️ Failed to load Excel file: {e}")
    sys.exit(1)


# --------------------------------------------
# Barcode Generator (Fixed for 999… SKUs)
# --------------------------------------------
def generate_barcode(sku: str) -> str:
    """
    Generates a Code128-B barcode PNG for the given SKU and returns the file path.
    Forces Code Set B to avoid digit dropping with Code Set C.
    """
    safe_sku = str(sku).strip()
    barcode_path = os.path.join(barcode_folder, safe_sku + ".png")

    if os.path.exists(barcode_path):
        return barcode_path

    # Force Code128-B to prevent auto numeric compaction issues
    code128 = barcode.get_barcode_class('code128')
    barcode_obj = code128(
        safe_sku,
        writer=ImageWriter(),
        charset='B'
    )

    barcode_obj.save(
        barcode_path.replace(".png", ""),
        options={"module_width": 0.22, "module_height": 8}
    )

    return barcode_path


# --------------------------------------------
# PDF Label Generator
# --------------------------------------------
def create_label_pdf(sku: str, description: str) -> str:
    pdf_path = os.path.join(downloads_folder, f"{sku}.pdf")
    c = canvas.Canvas(pdf_path, pagesize=(3 * inch, 1 * inch))

    # Load barcode image
    barcode_img_path = generate_barcode(sku)

    try:
        img = Image.open(barcode_img_path)
        img_width, img_height = img.size
        aspect = img_height / img_width

        target_width = 1.83 * inch
        target_height = target_width * aspect

        c.drawImage(
            barcode_img_path,
            x=0.6 * inch,
            y=0.4 * inch,
            width=target_width,
            height=target_height,
            preserveAspectRatio=True,
            anchor='c'
        )
    except Exception as e:
        st.error(f"⚠️ Failed to load barcode image: {e}")
        return None

    # Description text (multi-line wrap)
    if not description:
        description = ""

    c.setFont("Helvetica", 8)
    wrapped = wrap(description, width=34)

    text_y = 0.15 * inch
    for line in wrapped:
        text_width = c.stringWidth(line, "Helvetica", 8)
        c.drawString((3 * inch - text_width) / 2, text_y, line)
        text_y -= 0.12 * inch

    c.showPage()
    c.save()
    return pdf_path


# --------------------------------------------
# Streamlit UI
# --------------------------------------------
st.title("📦 SKU Barcode Generator")
st.write("Enter an SKU to generate a printable label.")

sku_input = st.text_input("Enter SKU:", "").strip()

# --------------------------------------------
# Button Logic
# --------------------------------------------
if st.button("Generate Label"):
    if not sku_input:
        st.warning("⚠️ Please enter a valid SKU.")
        st.stop()

    match = df[df["SKU"] == sku_input]

    if match.empty:
        st.warning("⚠️ SKU not found. Please try another.")
        st.stop()

    description = str(match.iloc[0].get("Description", ""))

    pdf_path = create_label_pdf(sku_input, description)

    if not pdf_path or not os.path.exists(pdf_path):
        st.error("❌ Failed to generate label PDF.")
        st.stop()

    # Convert PDF → Base64 for in-browser printing
    with open(pdf_path, "rb") as f:
        pdf_bytes = f.read()
        base64_pdf = base64.b64encode(pdf_bytes).decode("utf-8")

    print_button_html = f"""
    <script>
    function openPDF() {{
        const pdfData = atob("{base64_pdf}");
        const byteArray = new Uint8Array(pdfData.length);
        for (let i = 0; i < pdfData.length; i++) {{
            byteArray[i] = pdfData.charCodeAt(i);
        }}
        const blob = new Blob([byteArray], {{ type: 'application/pdf' }});
        const blobUrl = URL.createObjectURL(blob);
        const printWindow = window.open(blobUrl);
        printWindow.onload = function() {{
            setTimeout(() => {{
                printWindow.focus();
                printWindow.print();
            }}, 600);
        }};
    }}
    </script>
    <button onclick="openPDF()">🖨️ Print Label</button>
    """

    st.success("✅ Label generated. Click below to print.")
    components.html(print_button_html, height=100)
