import streamlit as st
from docx import Document
from pptx import Presentation
from pptx.util import Inches
import os

# Function to extract text from a Word document
def extract_text_from_docx(file):
    doc = Document(file)
    full_text = []
    for para in doc.paragraphs:
        full_text.append(para.text)
    return "\n".join(full_text)

# Function to create a PPT from the extracted text
def create_ppt(text, output_file):
    prs = Presentation()
    slide_layout = prs.slide_layouts[1]  # Title and Content layout

    # Split text into slides (e.g., by paragraphs or sections)
    slides_content = text.split("\n\n")  # Adjust based on your document structure

    for content in slides_content:
        slide = prs.slides.add_slide(slide_layout)
        title, body = slide.shapes.title, slide.shapes.placeholders[1]
        title.text = "Slide Title"  # Customize as needed
        body.text = content

    prs.save(output_file)

# Streamlit app
st.title("Word to PPT Converter")
st.write("Upload a Word document to generate a professional PPT.")

# File upload
uploaded_file = st.file_uploader("Upload a Word document", type=["docx"])

if uploaded_file is not None:
    # Extract text from the Word document
    text = extract_text_from_docx(uploaded_file)

    # Create a PPT
    ppt_file = "output_presentation.pptx"
    create_ppt(text, ppt_file)

    # Provide download links
    st.success("PPT generated successfully!")
    with open(ppt_file, "rb") as f:
        st.download_button("Download PPT", f, file_name=ppt_file)

    # Optional: Convert PPT to PDF
    # You can use libraries like `pdfkit` or `ReportLab` for this step