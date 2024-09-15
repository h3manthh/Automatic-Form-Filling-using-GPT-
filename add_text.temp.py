import fitz
import pandas as pd
import requests  # Added the missing import statement

def add_text_to_existing_pdf(input_pdf_path, answers, coordinates, output_pdf_path):
    """
    Fill in a form based on provided answers and their corresponding coordinates.
    
    :param input_pdf_path: Path to the source PDF.
    :param answers: Dictionary with form fields and their answers.
    :param coordinates: List of dictionaries with 'page', 'text', 'bbox' for each keyword.
    :param output_pdf_path: Path to save the modified PDF.
    """
    doc = fitz.open(input_pdf_path)
    modifications = []

    for key, value in answers.items():
        for coord in coordinates:
            if coord['text'].strip(':') == key:  # Match keyword without the colon
                page_number = coord['page']
                x1 = coord['bbox'][2]  # Right edge of the keyword's bbox
                y0, y1 = coord['bbox'][1], coord['bbox'][3]  # Bottom and top edge of the bbox
                y_mid = (y0 + y1) / 2  # Vertical midpoint for the text

                modifications.append({
                    'page': page_number,
                    'text': value,
                    'position': (x1 + 5, y_mid - 5),  # Adjust as needed for spacing
                    'size': 10  # Adjust font size as needed
                })
                break  # Move to the next keyword once matched

    for mod in modifications:
        page = doc.load_page(mod['page'])
        text = mod['text']
        position = mod['position']  # (x, y) coordinates
        size = mod.get('size', 11)  # Default font size is 11 if not specified
        
        # Add text to the page without wrap_text argument
        page.insert_text(position, text, fontsize=size)

    doc.save(output_pdf_path)

# Function to extract all text from a PDF
def extract_all_text(pdf_path):
    doc = fitz.open(pdf_path)
    all_text = ""
    for page in doc:
        page_text = page.get_text()
        all_text += page_text + "\n"
    return all_text

# Function to filter text by keywords and return formatted strings
def filter_text_by_keywords(extracted_text):
    keywords = [line.split(':', 1)[0] for line in extracted_text.split('\n') if ':' in line]
    keywords_Q = [line.split('?', 1)[0] for line in extracted_text.split('\n') if '?' in line]
    keywords_D = [line.split('.', 1)[0] for line in extracted_text.split('\n') if '.' in line]

    formatted_keywords = "\n".join(f'"{keyword}"' for keyword in keywords)
    formatted_keywords_Q = "\n".join(f'"{keywords_Q}"' for keywords_Q in keywords_Q)
    formatted_keywords_D = "\n".join(f'"{keywords_D}"' for keywords_D in keywords_D)

    return formatted_keywords, formatted_keywords_Q, formatted_keywords_D

# Specify the path to your PDF file
pdf_path = '/Users/hope/Desktop/Final Project/Data/BCU-application-form.pdf'

# Extract all text from the PDF
extracted_text = extract_all_text(pdf_path)

# Filter out lines containing keywords
formatted_keywords_str, formatted_keywords_str_Q, formatted_keywords_str_D = filter_text_by_keywords(extracted_text)

# Load the context data from a CSV file into a DataFrame
csv_file_path = '/Users/hope/Desktop/Final Project/Data/dummy.csv'
df = pd.read_csv(csv_file_path)

# Convert the DataFrame to a dictionary (assuming you want to use all data as context)
data_context_dict = df.to_dict(orient='records')[0]

# Convert the dictionary to a string context
data_context = '. '.join([f'"{key}": "{value}"' for key, value in data_context_dict.items()])

# Example usage (replace placeholders with your actual data and API key)
api_key = '.'
selected_model = 'gpt-3.5-turbo'
keywords = (formatted_keywords_str, formatted_keywords_str_Q, formatted_keywords_str_D)

# Function to call OpenAI GPT-3 API for chat completions
def call_chat_gpt(prompt, tokens, api_key, selected_model):
    url = "https://api.openai.com/v1/chat/completions"
    headers = {
        "Content-Type": "application/json",
        "Authorization": f"Bearer {api_key}"
    }
    payload = {
        "model": selected_model,
        "messages": [{"role": "user", "content": prompt}],
        "temperature": 0.7
    }
    response = requests.post(url, json=payload, headers=headers)
    return response.json()

# Function to clean up and extract the answer from OpenAI API response
def answer_cleanup(output):
    try:
        return output['choices'][0]['message']['content']
    except (KeyError, IndexError, TypeError):
        return "Error in processing the response."

# Function to generate answers for a list of keywords using the provided context
def generate_answers_for_keywords(context, keywords, api_key, selected_model):
    response = call_chat_gpt(context, 60, api_key, selected_model)
    cleaned_response = answer_cleanup(response)
    return {keyword: cleaned_response.strip() for keyword in keywords}

# Generate answers for the list of keywords
answers = generate_answers_for_keywords(data_context, keywords, api_key, selected_model)

# Add text to the PDF at the specified coordinates
input_pdf_path = '/Users/hope/Desktop/Final Project/Data/BCU-application-form.pdf'
output_pdf_path = '/Users/hope/Desktop/Final Project/Data/BCU-application-form-modified.pdf'

# Coordinates obtained using the extraction code
coordinates = [
    {'page': 0, 'text': 'Name:', 'bbox': [69.0, 142.5, 195.59999179840088, 156.0]},
    # Add more coordinates as needed
]

# Add text to the PDF and save the modified version
add_text_to_existing_pdf(input_pdf_path, answers, coordinates, output_pdf_path)
