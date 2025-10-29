import streamlit as st
import tempfile
import os
import pypandoc
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from docx import Document

st.set_page_config(page_title="PDF ↔ Word Converter", page_icon="📄", layout="centered")

st.title("📄 PDF ↔ Word Converter (Stable Layout Version)")
st.caption("Convert PDFs and Word documents — fully compatible with Streamlit Cloud or Hugging Face Spaces")

# --- Ensure pandoc is available ---
try:
    pypandoc.get_pandoc_version()
except OSError:
    with st.spinner("Installing Pandoc... please wait ⏳"):
        pypandoc.download_pandoc()

# UI
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
            st.info("Converting PDF → Word... please wait ⏳")
            # explicitly set input format
            pypandoc.convert_file(
                temp_file_path,
                "docx",
                format="pdf",
                outputfile=output_path,
                extra_args=["--standalone"]
            )
            output_file = output_path
            st.success("✅ Conversion successful — layout preserved where possible!")
        except Exception as e:
            st.error(f"Conversion failed: {e}")
            st.warning("Note: If your PDF is scanned or image-based, text extraction may not be possible.")

    elif option == "Word (.docx) ➜ PDF":
        output_path = temp_file_path + ".pdf"
        try:
            st.info("Converting Word → PDF... please wait ⏳")
            pypandoc.convert_file(
                temp_file_path,
                "pdf",
                format="docx",
                outputfile=output_path,
                extra_args=["--standalone"]
            )
            output_file = output_path
            st.success("✅ Conversion successful!")
        except Exception as e:
            st.error(f"Conversion failed: {e}")

    # download button
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
