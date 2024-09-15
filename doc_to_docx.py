import os
import subprocess

def convert_doc_to_docx(doc_file_path, docx_file_path):
    # Check if the file is a DOC file
    if not doc_file_path.lower().endswith('.doc'):
        print("The file is not a DOC file.")
        return

    # Convert DOC to DOCX using unoconv
    try:
        subprocess.run(['unoconv', '-f', 'docx', '-o', os.path.dirname(docx_file_path), doc_file_path])
        print(f"Conversion successful. DOCX file saved at: {docx_file_path}")
    except Exception as e:
        print(f"Error during conversion: {e}")

# Example usage
doc_file_path = '/Users/hope/Desktop/Final Project/Forms/PGT_Change_in_Terms_of_Study_Dec_2020.doc'
docx_file_path = '/Users/hope/Desktop/Final Project/Forms/converted/document.docx'

convert_doc_to_docx(doc_file_path, docx_file_path)
