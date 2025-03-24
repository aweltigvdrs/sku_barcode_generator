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
import time

# Setup
if getattr(sys, 'frozen', False):
    base_path = sys._MEIPASS
else:
    base_path = os.path.dirname(__file__)

downloads_folder = tempfile.gettempdir()
barcode_folder = os.path.join(base_path, "barcodes")
if not os.path.exists(barcode_folder):
    os.makedirs(barcode_folder)

excel_file = os.path.join(base_path, "sku_list.xlsx")
if not os.path.exists(excel_file):
    st.error(f"⚠️ Excel file not found at {excel_file}. Ensure it is present in the project folder.")
    sys.exit(1)

try:
    df = pd.read_excel(excel_file, dtype={"SKU": str})
except Exception as e:
    st.error(f"⚠️ Failed to load Excel file: {e}")
    sys.exit(1)

# Barcode generator
def generate_barcode(data):
    barcode_path = os.path.join(barcode_folder, data)
    final_path = barcode_path + ".png"
    if os.path.exists(final_path):
        return final_path
    code128 = barcode.get_barcode_class('code128')
    barcode_obj = code128(data, writer=ImageWriter())
    barcode_obj.save(barcode_path, options={"module_width": 0.25, "module_height": 8})
    return final_path

# Generate PDF with barcode and description
def create_label_pdf(sku, description):
    barcode_img_path = generate_barcode(sku)
    pdf_path = os.path.join(downloads_folder, f"{sku}.pdf")
    c = canvas.Canvas(pdf_path, pagesize=(3 * inch, 1 * inch))

    # Draw the barcode image
    try:
        img = Image.open(barcode_img_path)
        img_width, img_height = img.size
        img_ratio = img_height / img_width
        barcode_width = 1.83 * inch
        barcode_height = barcode_width * img_ratio
        c.drawImage(barcode_img_path, x=0.6 * inch, y=0.4 * inch, width=barcode_width, height=barcode_height, preserveAspectRatio=True)
    except Exception as e:
        st.error(f"⚠️ Failed to load barcode image: {e}")
        return None

    # Draw description
    c.setFont("Helvetica", 8)
    text = description[:34]
    text_width = c.stringWidth(text, "Helvetica", 8)
    c.drawString((3 * inch - text_width) / 2, 0.1 * inch, text)

    c.showPage()
    c.save()
    return pdf_path

# --- Streamlit UI ---
st.title("📦 SKU Barcode Generator (Linux-Compatible)")
st.write("Enter an SKU to generate a printable label.")

sku_input = st.text_input("Enter SKU:", "")

if st.button("Generate Label"):
    if sku_input.strip():
        match = df[df["SKU"] == sku_input]
        if match.empty:
            st.warning("⚠️ SKU not found. Please try another.")
        else:
            description = match.iloc[0]["Description"]
            time.sleep(1)
            pdf_path = create_label_pdf(sku_input, description)

            if pdf_path and os.path.exists(pdf_path):
                with open(pdf_path, "rb") as f:
                    base64_pdf = base64.b64encode(f.read()).decode("utf-8")

                pdf_viewer = f"""
                <iframe src="data:application/pdf;base64,{base64_pdf}" width="400" height="300"></iframe>
                <br>
                <button onclick="printPDF()">🖨️ Print Label</button>
                <script>
                function printPDF() {{
                    const iframe = document.createElement('iframe');
                    iframe.style.display = 'none';
                    iframe.src = "data:application/pdf;base64,{base64_pdf}";
                    document.body.appendChild(iframe);
                    iframe.onload = function() {{
                        setTimeout(() => {{
                            iframe.contentWindow.focus();
                            iframe.contentWindow.print();
                        }}, 500);
                    }};
                }}
                </script>
                """
                st.success("✅ Label ready for print below.")
                components.html(pdf_viewer, height=400)
            else:
                st.error("❌ Failed to generate label PDF.")
    else:
        st.warning("⚠️ Please enter a valid SKU.")
