import fitz  # Import the PyMuPDF library

def extract_all_text(pdf_path):
    # Open the PDF file
    doc = fitz.open(pdf_path)
    
    all_text = ""  # Initialize a variable to store all the extracted text

    # Iterate through each page in the PDF
    for page in doc:
        # Extract text from the current page
        page_text = page.get_text()
        all_text += page_text + "\n"  # Append the text from this page, add a newline as separator

    return all_text

# Specify the path to your PDF file
pdf_path = '/Users/hope/Desktop/Final Project/Data/BCU-application-form.pdf'

# Extract all text
extracted_text = extract_all_text(pdf_path)

# Optionally, print the extracted text to verify
print(extracted_text)
