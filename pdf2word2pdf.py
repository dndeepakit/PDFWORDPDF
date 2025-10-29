import streamlit as st
import pdfplumber
from docx import Document
from io import BytesIO
import fitz  # PyMuPDF

st.set_page_config(page_title="PDF ⇄ Word Converter", page_icon="📄", layout="centered")

def convert_pdf_to_word(pdf_bytes):
    """Convert PDF to Word and preserve text + tables layout."""
    output = BytesIO()
    doc = Document()

    with pdfplumber.open(BytesIO(pdf_bytes)) as pdf:
        for page_num, page in enumerate(pdf.pages, start=1):
            doc.add_heading(f"Page {page_num}", level=2)

            # Extract tables first
            tables = page.extract_tables()
            if tables:
                for table in tables:
                    rows, cols = len(table), len(table[0])
                    table_obj = doc.add_table(rows=rows, cols=cols)
                    for i, row in enumerate(table):
                        for j, cell in enumerate(row):
                            table_obj.cell(i, j).text = str(cell) if cell else ""
                    doc.add_paragraph("")  # spacing

            # Extract text
            text = page.extract_text(x_tolerance=1, y_tolerance=2)
            if not text:
                # fallback to PyMuPDF
                with fitz.open(stream=pdf_bytes, filetype="pdf") as doc_fitz:
                    text = doc_fitz.load_page(page_num - 1).get_text("text")
            if text:
                doc.add_paragraph(text)
            doc.add_page_break()

    doc.save(output)
    return output.getvalue()

def main():
    st.title("📄 PDF ⇄ Word Converter")
    st.write("Convert your PDF files into editable Word documents (layout & tables preserved).")

    uploaded_file = st.file_uploader("Upload a PDF file", type=["pdf"])

    if uploaded_file is not None:
        pdf_bytes = uploaded_file.read()
        with st.spinner("Converting... please wait ⏳"):
            try:
                word_data = convert_pdf_to_word(pdf_bytes)
                st.success("✅ Conversion successful!")
                st.download_button(
                    label="📥 Download Word File",
                    data=word_data,
                    file_name="converted.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
            except Exception as e:
                st.error(f"❌ Conversion failed: {e}")

    st.caption("Built with ❤️ using Streamlit + PyMuPDF + pdfplumber")

if __name__ == "__main__":
    main()
