import streamlit as st
from pdf2docx import Converter
import tempfile
import os

st.set_page_config(page_title="PDF ↔ Word Converter", page_icon="📄", layout="centered")

st.title("📄 PDF ↔ Word Converter (Free Cloud Version)")
st.write("Convert between PDF and Word online — no installation needed!")

option = st.radio("Choose conversion type:", ["PDF → Word", "Word → PDF"])

uploaded_file = st.file_uploader("Upload your file", type=["pdf", "docx"])

if uploaded_file:
    with tempfile.NamedTemporaryFile(delete=False, suffix=f".{uploaded_file.name.split('.')[-1]}") as tmp:
        tmp.write(uploaded_file.read())
        input_path = tmp.name

    if option == "PDF → Word" and uploaded_file.name.endswith(".pdf"):
        output_path = input_path.replace(".pdf", ".docx")
        try:
            st.info("⏳ Converting... please wait.")
            cv = Converter(input_path)
            cv.convert(output_path, start=0, end=None)
            cv.close()
            with open(output_path, "rb") as f:
                st.download_button("⬇️ Download Converted Word File", f, file_name="converted.docx")
            st.success("✅ Conversion completed successfully!")
        except Exception as e:
            st.error(f"Conversion failed: {e}")

    elif option == "Word → PDF" and uploaded_file.name.endswith(".docx"):
        try:
            from fpdf import FPDF
            from docx import Document

            doc = Document(input_path)
            pdf = FPDF()
            pdf.add_page()
            pdf.set_auto_page_break(auto=True, margin=15)
            pdf.set_font("Arial", size=12)

            for para in doc.paragraphs:
                pdf.multi_cell(0, 10, para.text)

            output_path = input_path.replace(".docx", ".pdf")
            pdf.output(output_path)

            with open(output_path, "rb") as f:
                st.download_button("⬇️ Download Converted PDF", f, file_name="converted.pdf")
            st.success("✅ Conversion completed successfully!")
        except Exception as e:
            st.error(f"Conversion failed: {e}")
    else:
        st.warning("⚠️ Please upload the correct file type for your selected conversion.")
