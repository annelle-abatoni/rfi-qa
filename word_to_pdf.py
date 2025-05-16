import os
import comtypes.client

def convert_word_to_pdf(docx_path, pdf_path):
    word = comtypes.client.CreateObject('Word.Application')
    doc = word.Documents.Open(docx_path)
    doc.SaveAs(pdf_path, FileFormat=17)  # 17 is the format for PDF
    doc.Close()
    word.Quit()

def process_word_files(root_dir):
    for foldername, subfolders, filenames in os.walk(root_dir):
        for filename in filenames:
            if filename.lower().endswith('.docx') and not filename.startswith('~$'):
                full_docx_path = os.path.join(foldername, filename)
                
                new_name = filename[3:] if len(filename) > 3 else filename
                new_pdf_name = os.path.splitext(new_name)[0] + '.pdf'
                full_pdf_path = os.path.join(foldername, new_pdf_name)

                try:
                    print(f"Converting: {filename} -> {new_pdf_name}")
                    convert_word_to_pdf(full_docx_path, full_pdf_path)
                except Exception as e:
                    print(f"Failed to convert {filename}: {e}")

# === Set the directory you want to scan ===
directory_path = r"C:\path\to\your\word\docs"

process_word_files(directory_path)
