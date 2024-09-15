import pytesseract
from pdf2image import convert_from_path

def extract_text_from_pdf(pdf_path):
    # Convert PDF to images
    images = convert_from_path(pdf_path)

    # Extract text from each image using OCR
    extracted_text = ""
    for image in images:
        text = pytesseract.image_to_string(image, lang='eng')
        extracted_text += text

    return extracted_text

# Path to the PDF file
pdf_path = "/Users/hope/Desktop/Final Project/24MARCH/Forms/CDC-Small-Grants-Application-Form-final-1.pdf"

# Extract text from the PDF
text = extract_text_from_pdf(pdf_path)

# Print the extracted text
print(text)

