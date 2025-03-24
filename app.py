import os
import sys
import pandas as pd
import barcode
import tempfile
from barcode.writer import ImageWriter
from PIL import ImageFont
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
import shutil
import streamlit as st
import time

# Set up paths
if getattr(sys, 'frozen', False):
    base_path = sys._MEIPASS  # PyInstaller temp folder
else:
    base_path = os.path.dirname(__file__)

downloads_folder = tempfile.gettempdir()
barcode_folder = os.path.join(base_path, "barcodes")  # Save barcodes persistently

# Ensure barcode folder exists
if not os.path.exists(barcode_folder):
    os.makedirs(barcode_folder)

excel_file = os.path.join(base_path, "sku_list.xlsx")

# Load SKU list from Excel
if not os.path.exists(excel_file):
    st.error(f"⚠️ Excel file not found at {excel_file}. Ensure it is present in the project folder.")
    sys.exit(1)

try:
    df = pd.read_excel(excel_file, dtype={"SKU": str})
except Exception as e:
    st.error(f"⚠️ Failed to load Excel file: {e}")
    sys.exit(1)

# Function to generate barcode
def generate_barcode(data):
    # Define correct barcode file path (without .png)
    barcode_path = os.path.join(barcode_folder, data)  # Remove ".png"

    # Debugging: Print the expected barcode path
    print(f"📌 Expected barcode path: {barcode_path}.png")

    # Skip generation if barcode already exists
    barcode_final_path = barcode_path + ".png"
    if os.path.exists(barcode_final_path):  # Check if actual file exists
        print(f"✅ Barcode already exists at: {barcode_final_path}")
        return barcode_final_path

    # Generate barcode
    code128 = barcode.get_barcode_class('code128')
    barcode_obj = code128(data, writer=ImageWriter())

    try:
        # Save barcode image (this function automatically adds .png)
        barcode_obj.save(barcode_path, options={"module_width": 0.25, "module_height": 8})

        # Double-check if barcode exists after saving
        if not os.path.exists(barcode_final_path):
            print(f"❌ ERROR: Expected barcode at {barcode_final_path} but it was NOT found!")
            raise FileNotFoundError(f"🚨 Barcode file {barcode_final_path} was not found after saving!")

        print(f"✅ Barcode successfully generated at: {barcode_final_path}")
        return barcode_final_path

    except Exception as e:
        print(f"⚠️ Error in generating barcode: {e}")
        sys.exit(1)



# Function to create Word document
def create_word_doc(sku, description):
    doc = Document()
    section = doc.sections[0]
    section.page_width = Inches(3)
    section.page_height = Inches(1)
    section.top_margin = Inches(0.05)
    section.bottom_margin = Inches(0.05)
    section.left_margin = Inches(0.05)
    section.right_margin = Inches(0.05)

    img_filename = generate_barcode(sku)

    # Ensure the barcode image exists before proceeding
    if not os.path.exists(img_filename):
        st.error(f"⚠️ Barcode file {img_filename} is missing. Generation failed.")
        sys.exit(1)

    # Add barcode image
    barcode_para = doc.add_paragraph()
    barcode_run = barcode_para.add_run()
    barcode_run.add_picture(img_filename, width=Inches(1.83))
    barcode_para.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    barcode_para.paragraph_format.space_before = Pt(0)
    barcode_para.paragraph_format.space_after = Pt(0)

    # Add description text
    desc_para = doc.add_paragraph(description[:34])
    desc_para.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    desc_run = desc_para.runs[0]
    desc_run.font.size = Pt(11)
    desc_run.font.name = 'Arial'
    desc_para.paragraph_format.space_before = Pt(0)
    desc_para.paragraph_format.space_after = Pt(0)

    # Save the document
    doc_filename = os.path.join(downloads_folder, f"{sku}.docx")
    doc.save(doc_filename)

    return doc_filename

# --- Streamlit UI ---
st.title("📦 SKU Barcode Generator")
st.write("Enter an SKU to generate a barcode and label file.")

sku_input = st.text_input("Enter SKU:", "")

if st.button("Generate Label"):
    if sku_input.strip():
        match = df[df["SKU"] == sku_input]
        if match.empty:
            st.warning("⚠️ SKU not found. Please try another.")
        else:
            description = match.iloc[0]["Description"]
            
            # Introduce a small delay for file system sync
            time.sleep(1)

            doc_path = create_word_doc(sku_input, description)
            st.success(f"✅ Label generated: {doc_path}")
            st.download_button(
                label="📥 Download Label File",
                data=open(doc_path, "rb"),
                file_name=f"{sku_input}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
    else:
        st.warning("⚠️ Please enter a valid SKU.")
