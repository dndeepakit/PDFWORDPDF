import streamlit as st
import fitz  # PyMuPDF
from docx import Document
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
import tempfile
import os

st.set_page_config(page_title="PDF ↔ Word Converter", page_icon="📄", layout="centered")

st.title("📄 PDF ↔ Word Converter")
st.write("Convert PDF ↔ Word documents easily and securely in your browser.")

option = st.radio("Select Conversion Type:", ("PDF ➜ Word (.docx)", "Word (.docx) ➜ PDF"))

uploaded_file = st.file_uploader("Upload your file", type=["pdf", "docx"])

if uploaded_file:
    with tempfile.NamedTemporaryFile(delete=False) as temp_file:
        temp_file.write(uploaded_file.read())
        temp_file_path = temp_file.name

    output_file = None

    if option == "PDF ➜ Word (.docx)":
        output_path = temp_file_path + ".docx"
        try:
            doc = fitz.open(temp_file_path)
            word_doc = Document()
            for page_num, page in enumerate(doc, start=1):
                text = page.get_text("text")
                word_doc.add_paragraph(text)
            word_doc.save(output_path)
            doc.close()
            output_file = output_path
            st.success("✅ PDF successfully converted to Word!")
        except Exception as e:
            st.error(f"Conversion failed: {e}")

    elif option == "Word (.docx) ➜ PDF":
        output_path = temp_file_path + ".pdf"
        try:
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
            st.success("✅ Word successfully converted to PDF!")
        except Exception as e:
            st.error(f"Conversion failed: {e}")

    if output_file and os.path.exists(output_file):
        with open(output_file, "rb") as f:
            st.download_button(
                label="⬇️ Download Converted File",
                data=f,
                file_name=os.path.basename(output_file),
                mime="application/octet-stream"
            )

    # Clean up temp files
    os.remove(temp_file_path)
    if output_file and os.path.exists(output_file):
        os.remove(output_file)
