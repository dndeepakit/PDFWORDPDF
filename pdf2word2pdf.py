import streamlit as st
from pdf2docx import Converter
from fpdf import FPDF
from docx import Document
import tempfile
import os
import traceback

st.set_page_config(page_title="PDF ↔ Word Converter", page_icon="📄", layout="centered")

st.title("📄 PDF ↔ Word Converter (Streamlit Cloud)")
st.caption("Convert between PDF and Word — no installation, 100% free!")

option = st.radio("Choose conversion type:", ["PDF → Word", "Word → PDF"])
uploaded_file = st.file_uploader("Upload your file", type=["pdf", "docx"])

if uploaded_file:
    suffix = os.path.splitext(uploaded_file.name)[1].lower()
    with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
        tmp.write(uploaded_file.read())
        input_path = tmp.name

    try:
        if option == "PDF → Word" and suffix == ".pdf":
            output_path = input_path.replace(".pdf", ".docx")
            st.info("⏳ Converting your PDF to Word... please wait.")
            try:
                cv = Converter(input_path)
                cv.convert(output_path, start=0, end=None)
                cv.close()
            except Exception as e:
                st.warning("⚠️ PDF2DOCX failed, using backup converter...")
                try:
                    import fitz  # PyMuPDF
                    doc = Document()
                    pdf = fitz.open(input_path)
                    for page in pdf:
                        text = page.get_text("text")
                        doc.add_paragraph(text)
                    doc.save(output_path)
                except Exception as e2:
                    st.error(f"Conversion failed in both methods: {e2}")
                    st.text(traceback.format_exc())
                    st.stop()

            with open(output_path, "rb") as f:
                st.download_button("⬇️ Download Converted Word File", f, file_name="converted.docx")
            st.success("✅ Done!")

        elif option == "Word → PDF" and suffix == ".docx":
            output_path = input_path.replace(".docx", ".pdf")
            st.info("⏳ Converting your Word to PDF...")
            doc = Document(input_path)
            pdf = FPDF()
            pdf.add_page()
            pdf.set_font("Arial", size=12)
            for para in doc.paragraphs:
                pdf.multi_cell(0, 10, para.text)
            pdf.output(output_path)

            with open(output_path, "rb") as f:
                st.download_button("⬇️ Download Converted PDF", f, file_name="converted.pdf")
            st.success("✅ Done!")

        else:
            st.warning("⚠️ Please upload a matching file type for your selection.")
    except Exception as e:
        st.error(f"Conversion failed: {e}")
        st.text(traceback.format_exc())
