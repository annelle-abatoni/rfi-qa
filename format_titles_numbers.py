import os
import re
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn

def is_number_with_comma(text):
    return re.sub(r'(?<=\d),(?=\d{2}\b)', '.', text)

def apply_arial_font_to_run(run):
    run.font.name = 'Arial'
    run._element.rPr.rFonts.set(qn('w:eastAsia'), 'Arial')

def process_document(input_path):
    if not os.path.exists(input_path):
        print(f"❌ File not found: {input_path}")
        return

    doc = Document(input_path)

    page_counter = 1
    para_counter = 0
    title_number = 1
    processed_page_breaks = 0

    for para in doc.paragraphs:
        # Convert all font to Arial
        for run in para.runs:
            apply_arial_font_to_run(run)
            # Replace numbers like 3,45 with 3.45
            run.text = is_number_with_comma(run.text)

        # Detect and count page breaks
        if any(run.text == '\f' for run in para.runs):
            processed_page_breaks += 1

        if processed_page_breaks >= 3:  # Page 4 onwards
            text = para.text.strip()
            if text and para.style.name.startswith("Heading"):
                # Apply title formatting
                para.text = f"{title_number}. {text}"
                para.style.font.name = 'Arial'
                para.runs[0].font.bold = True
                para.runs[0].font.size = Pt(12)
                apply_arial_font_to_run(para.runs[0])
                title_number += 1

    output_path = os.path.splitext(input_path)[0] + "_formatted.docx"
    doc.save(output_path)
    print(f"✅ Formatted document saved as: {output_path}")

# --- Run the script ---
input_path = input("Enter full path to your .docx file: ").strip('" ')
process_document(input_path)