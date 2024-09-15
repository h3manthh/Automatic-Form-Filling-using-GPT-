import pandas as pd
import openai
import requests
import fitz
import pandas as pd
import openai
import requests
import fitz
from docx import Document

# Function to extract text from a PDF file
def extract_text_from_pdf(pdf_path):
    doc = fitz.open(pdf_path)
    all_text = [page.get_text() for page in doc]
    return all_text

# Function to extract text from a Word file
def extract_text_from_docx(docx_path):
    doc = Document(docx_path)
    all_text = [paragraph.text for paragraph in doc.paragraphs]
    return all_text

# Function to filter text by keywords
def filter_text_by_keywords(extracted_text):
    keywords = [line.split(':', 1)[0] for line in extracted_text.split('\n') if ':' in line]  # using :
    keywords_Q = [line.split('?', 1)[0] for line in extracted_text.split('\n') if '?' in line]  # using ?
    keywords_D = [line.split('.', 1)[0] for line in extracted_text.split('\n') if '.' in line]  # using .
    formatted_keywords = {f'"{keyword}"': None for keyword in keywords}
    formatted_keywords_Q = {f'"{keyword}"': None for keyword in keywords_Q}
    formatted_keywords_D = {f'"{keyword}"': None for keyword in keywords_D}
    return formatted_keywords, formatted_keywords_Q, formatted_keywords_D

# Function to extract text from a file (PDF or Word)
def extract_text_from_file(file_path):
    if file_path.lower().endswith('.pdf'):
        return extract_text_from_pdf(file_path)
    elif file_path.lower().endswith('.docx'):
        return extract_text_from_docx(file_path)
    else:
        raise ValueError("Unsupported file type. Only PDF and Word files are supported.")

# Specify the path to your file (PDF or Word)
file_path = '/Users/hope/Desktop/Final Project/Forms/PGT_Readmission_form_Dec_2020.docx'

# Extract text from the entire file using the appropriate function
extracted_text_from_file = extract_text_from_file(file_path)

# Filter out lines containing keywords
formatted_keywords_str, formatted_keywords_str_Q, formatted_keywords_str_D = filter_text_by_keywords('\n'.join(extracted_text_from_file))


# Function to extract text from a PDF file
def extract_text_from_pdf(pdf_path):
    doc = fitz.open(pdf_path)
    all_text = [page.get_text() for page in doc]
    return all_text

# Function to filter text by keywords
def filter_text_by_keywords(extracted_text):
    keywords = [line.split(':', 1)[0] for line in extracted_text.split('\n') if ':' in line]  # using :
    keywords_Q = [line.split('?', 1)[0] for line in extracted_text.split('\n') if '?' in line]  # using ?
    keywords_D = [line.split('.', 1)[0] for line in extracted_text.split('\n') if '.' in line]  # using .
    formatted_keywords = {f'"{keyword}"': None for keyword in keywords}
    formatted_keywords_Q = {f'"{keyword}"': None for keyword in keywords_Q}
    formatted_keywords_D = {f'"{keyword}"': None for keyword in keywords_D}
    return formatted_keywords, formatted_keywords_Q, formatted_keywords_D

# Specify the path to your PDF file
pdf_path = '/Users/hope/Desktop/Final Project/Forms/PGT_Readmission_form_Dec_2020.docx'

# Extract text from the entire PDF using PyMuPDF
extracted_text_from_pdf = extract_text_from_pdf(pdf_path)

# Filter out lines containing keywords
formatted_keywords_str, formatted_keywords_str_Q, formatted_keywords_str_D = filter_text_by_keywords('\n'.join(extracted_text_from_pdf))

# Load the context data from a CSV file into a DataFrame
csv_file_path = '/Users/hope/Desktop/Final Project/Data/dummy_data.csv'
df = pd.read_csv(csv_file_path)

# Convert the DataFrame to a dictionary (assuming you want to use all data as context)
data_context_dict = df.to_dict(orient='records')[0]  # Taking the first row as an example

# Convert the dictionary to a string context
data_context = '. '.join([f'"{key}": {value}' for key, value in data_context_dict.items()])

# Example usage (replace placeholders with your actual data and API key)
api_key = 'api key'
selected_model = 'gpt-3.5-turbo'

# Concatenate the keyword lists into a single list
keywords = list(formatted_keywords_str.keys()) + list(formatted_keywords_str_Q.keys()) + list(formatted_keywords_str_D.keys())

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
    return response.json()  # Ensures correct API response capture

# Function to clean up and extract the answer from OpenAI API response
def answer_cleanup(output):
    try:
        return output['choices'][0]['message']['content']
    except (KeyError, IndexError, TypeError):
        return "Error in processing the response."

# Function to generate answers for a list of keywords using the provided context
def generate_answers_for_keywords(context, keywords, api_key, selected_model):
    answers = {}

    for keyword in keywords:
        prompt = f"Based on the following details: {context}, what is the answer for '{keyword}'? Keep the response in mostly 1 word or few words, maximum of 5 words."
        response = call_chat_gpt(prompt, 60, api_key, selected_model)  # Example token count; adjust as needed
        cleaned_response = answer_cleanup(response)
        answers[keyword] = cleaned_response.strip()

    return answers

# Generate answers for the list of keywords
answers = generate_answers_for_keywords(data_context, keywords, api_key, selected_model)

# Print the generated answers
for keyword, answer in answers.items():
    print(f"{keyword}: {answer}\n")
