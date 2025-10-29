import streamlit as st
import pdfplumber
from docx import Document
from io import BytesIO
import fitz  # PyMuPDF

st.set_page_config(page_title="PDF ⇄ Word Converter", page_icon="📄", layout="centered")

def convert_pdf_to_word(pdf_bytes):
    """Extract text, tables and layout from PDF and create a formatted Word doc."""
    output = BytesIO()
    doc = Document()

    with pdfplumber.open(BytesIO(pdf_bytes)) as pdf:
        for page_num, page in enumerate(pdf.pages, start=1):
            # Add page header
            doc.add_heading(f"Page {page_num}", level=2)

            # Try to extract tables first
            tables = page.extract_tables()
            if tables:
                for table in tables:
                    rows = len(table)
                    cols = len(table[0])
                    t = doc.add_table(rows=rows, cols=cols)
                    for i, row in enumerate(table):
                        for j, cell in enumerate(row):
                            t.cell(i, j).text = str(cell) if cell else ""
                    doc.add_paragraph("")  # spacing after table

            # Extract text blocks using PyMuPDF for order accuracy
            page_text = page.extract_text(x_tolerance=1, y_tolerance=2)
            if page_text:
                doc.add_paragraph(page_text)
            else:
                # fallback: low-level text extraction
                with fitz.open(stream=pdf_bytes, filetype="pdf") as doc_fitz:
                    text = doc_fitz.load_page(page_num - 1).get_text("text")
                    doc.add_paragraph(text)

            doc.add_page_break()

    doc.save(output)
    return output.getvalue()


def main():
    st.title("📄 PDF ⇄ Word Converter")
    st.write("Convert your PDF files into editable Word documents with preserved layout and tables.")

    uploaded_file = st.file_uploader("Upload a PDF file", type=["pdf"])

    if uploaded_file is not None:
        pdf_bytes = uploaded_file.read()
        with st.spinner("Converting... please wait ⏳"):
            try:
                docx_data = convert_pdf_to_word(pdf_bytes)
                st.success("✅ Conversion successful!")
                st.download_button(
                    label="📥 Download Word File",
                    data=docx_data,
                    file_name="converted.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
            except Exception as e:
                st.error(f"❌ Conversion failed: {e}")

    st.markdown("---")
    st.caption("Built with ❤️ using Streamlit, pdfplumber, and PyMuPDF")


if __name__ == "__main__":
    main()
