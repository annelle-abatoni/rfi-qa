import os
import comtypes.client

def convert_docx_to_pdf(docx_path, pdf_path):
    try:
        word = comtypes.client.CreateObject('Word.Application')
        word.Visible = False
        doc = word.Documents.Open(docx_path)
        doc.SaveAs(pdf_path, FileFormat=17)  # 17 = wdFormatPDF
        doc.Close()
        word.Quit()
        print(f"✅ PDF saved at: {pdf_path}")
    except Exception as e:
        print(f"❌ Failed to convert {docx_path}: {e}")

def process_all_docx_files(root_folder):
    for foldername, _, filenames in os.walk(root_folder):
        for filename in filenames:
            if filename.lower().endswith(".docx") and not filename.startswith("~$"):
                full_docx_path = os.path.join(foldername, filename)
                
                # Remove first 3 characters from filename
                trimmed_name = filename[5:] if len(filename) > 5 else filename
                trimmed_basename = os.path.splitext(trimmed_name)[0]
                pdf_name = trimmed_basename + ".pdf"
                full_pdf_path = os.path.join(foldername, pdf_name)

                convert_docx_to_pdf(full_docx_path, full_pdf_path)

# --- Run the script ---
root = input("Enter the root folder to search for .docx files: ").strip('" ')
if os.path.isdir(root):
    process_all_docx_files(root)
else:
    print("❌ Invalid directory.")
