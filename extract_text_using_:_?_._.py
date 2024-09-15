import fitz

def extract_text_from_pdf(pdf_path):
    # Open the PDF file
    doc = fitz.open(pdf_path)

    all_text = []  # Initialize a list to store extracted text

    # Iterate through each page in the PDF
    for page_number in range(doc.page_count):
        # Extract text from the current page
        page = doc[page_number]
        page_text = page.get_text()

        all_text.append(page_text)

    return all_text

def filter_text_by_keywords(extracted_text):

    # Filter out lines containing keywords
    keywords = [line.split(':', 1)[0] for line in extracted_text.split('\n') if ':' in line] #using :
    keywords_Q = [line.split('?', 1)[0] for line in extracted_text.split('\n') if '?' in line] #using ?
    keywords_D = [line.split('.', 1)[0] for line in extracted_text.split('\n') if '.' in line] #using .

    # Convert keywords to the desired format

    formatted_keywords = {f'"{keyword}"': None for keyword in keywords}
    formatted_keywords_Q = {f'"{keywords_Q}"': None for keywords_Q in keywords_Q}
    formatted_keywords_D = {f'"{keywords_D}"': None for keywords_D in keywords_D}


    # Print the formatted keywords
    formatted_keywords_str ="\n".join(formatted_keywords.keys())
    print((formatted_keywords_str))
    formatted_keywords_str_Q ="\n".join(formatted_keywords_Q.keys())
    print((formatted_keywords_str_Q))
    formatted_keywords_str_D ="\n".join(formatted_keywords_D.keys())
    print((formatted_keywords_str_D))


# Specify the path to your PDF file
pdf_path = '/Users/hope/Desktop/Final Project/Data/BCU-application-form.pdf'

# Extract text from the entire PDF using PyMuPDF
extracted_text_from_pdf = extract_text_from_pdf(pdf_path)

# Filter out lines containing keywords
filter_text_by_keywords('\n'.join(extracted_text_from_pdf))
