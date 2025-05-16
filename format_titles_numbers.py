import os
import comtypes.client

def convert_docx_to_pdf(input_path):
    if not os.path.exists(input_path):
        print("❌ File does not exist.")
        return

    folder = os.path.dirname(input_path)
    filename = os.path.basename(input_path)

    if not filename.lower().endswith('.docx'):
        print("❌ Not a .docx file.")
        return

    # Remove first 3 characters from the filename
    trimmed_name = filename[5:] if len(filename) > 5 else filename
    pdf_filename = os.path.splitext(trimmed_name)[0] + '.pdf'
    output_path = os.path.join(folder, pdf_filename)

    try:
        word = comtypes.client.CreateObject('Word.Application')
        doc = word.Documents.Open(input_path)
        doc.SaveAs(output_path, FileFormat=17)  # 17 is for PDF
        doc.Close()
        word.Quit()
        print(f"✅ PDF saved at: {output_path}")
    except Exception as e:
        print(f"❌ Failed to convert: {e}")

# --- Run the script ---
input_path = input("Enter full path to your .docx file: ").strip('" ')
convert_docx_to_pdf(input_path)
