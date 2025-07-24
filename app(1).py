import streamlit as st
import pytesseract
from pdf2image import convert_from_bytes
import pandas as pd
import re
from PIL import ImageEnhance
import io

st.title("PDF to Excel Extractor - Fabric Info")

uploaded_file = st.file_uploader("Upload a PDF", type=["pdf"])

if uploaded_file:
    # Convert PDF to images
    images = convert_from_bytes(uploaded_file.read(), dpi=300)

    entries = []

    for image in images:
        # Enhance image contrast and convert to grayscale
        img = ImageEnhance.Contrast(image.convert("L")).enhance(2.0)
        ocr_text = pytesseract.image_to_string(img)

        # Line-by-line parsing
        lines = [line.strip() for line in ocr_text.splitlines() if line.strip()]
        current_entry = {}

        for line in lines:
            if "PRINT COLOR" in line.upper():
                match = re.search(r"PRINT COLOR\s+(.+)", line, re.IGNORECASE)
                if match:
                    current_entry["Print Color"] = match.group(1).split('|')[0].split('MANAOLA BUTTON')[0].strip()

            if "BUTTON" in line.upper():
                match = re.search(r"BUTTON\s+(.+)", line, re.IGNORECASE)
                if match:
                    current_entry["Button Color"] = match.group(1).strip()

            if "FABRIC COLOR" in line.upper():
                match = re.search(r"FABRIC COLOR\s+(.+)", line, re.IGNORECASE)
                if match:
                    current_entry["Fabric Color"] = match.group(1).split('PATTERN')[0].strip()

            if "PATTERN" in line.upper():
                match = re.search(r"PATTERN\s+(.+)", line, re.IGNORECASE)
                if match:
                    current_entry["Pattern"] = match.group(1).strip()

            if len(current_entry) == 4:
                entries.append(current_entry)
                current_entry = {}

    if entries:
        df = pd.DataFrame(entries)
        st.dataframe(df)

        # Convert to Excel and store in-memory
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False)
        output.seek(0)

        st.download_button(
            label="Download Excel",
            data=output,
            file_name="fabric_info.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.warning("No fabric information detected. Please check the PDF content.")