import streamlit as st
import tempfile
import os
import pdfplumber
from docx import Document
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter

st.set_page_config(page_title="PDF ↔ Word Converter", page_icon="📄", layout="centered")

st.title("📄 PDF ↔ Word Converter")
st.caption("Convert your PDF to Word (layout preserved) and vice versa. Works fully offline.")

option = st.radio("Select Conversion Type:", ("PDF ➜ Word (.docx)", "Word (.docx) ➜ PDF"))

uploaded_file = st.file_uploader("Upload your file", type=["pdf", "docx"])

if uploaded_file:
    with tempfile.NamedTemporaryFile(delete=False) as temp_file:
        temp_file.write(uploaded_file.read())
        temp_file_path = temp_file.name

    output_file = None

    # -------- PDF to Word --------
    if option == "PDF ➜ Word (.docx)":
        output_path = temp_file_path + ".docx"
        try:
            st.info("Converting PDF → Word... please wait ⏳")

            doc = Document()
            with pdfplumber.open(temp_file_path) as pdf:
                for page in pdf.pages:
                    text = page.extract_text()
                    if text:
                        doc.add_paragraph(text)
                        doc.add_page_break()
                    else:
                        st.warning("⚠️ Some pages had no extractable text (may contain images only).")

            doc.save(output_path)
            output_file = output_path
            st.success("✅ Conversion successful!")

        except Exception as e:
            st.error(f"Conversion failed: {e}")

    # -------- Word to PDF --------
    elif option == "Word (.docx) ➜ PDF":
        output_path = temp_file_path + ".pdf"
        try:
            st.info("Converting Word → PDF... please wait ⏳")
            from docx import Document
            doc = Document(temp_file_path)
            pdf = canvas.Canvas(output_path, pagesize=letter)
            width, height = letter
            y = height - 50
            for para in doc.paragraphs:
                text = para.text
                pdf.drawString(50, y, text)
                y -= 15
                if y < 50:
                    pdf.showPage()
                    y = height - 50
            pdf.save()
            output_file = output_path
            st.success("✅ Conversion successful!")

        except Exception as e:
            st.error(f"Conversion failed: {e}")

    # -------- Download --------
    if output_file and os.path.exists(output_file):
        with open(output_file, "rb") as f:
            st.download_button(
                label="⬇️ Download Converted File",
                data=f,
                file_name=os.path.basename(output_file),
                mime="application/octet-stream"
            )

    # cleanup
    os.remove(temp_file_path)
    if output_file and os.path.exists(output_file):
        os.remove(output_file)
