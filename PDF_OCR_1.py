import fitz
import pandas as pd
import openai
import requests
import pytesseract
from PIL import Image
import io

# Function to extract all text from a PDF
def extract_all_text(pdf_path):
    doc = fitz.open(pdf_path)
    all_text = ""
    for page in doc:
        page_text = page.get_text()
        all_text += page_text + "\n"
    return all_text

# Function to load CSV data into a DataFrame
def load_csv(csv_file_path):
    df = pd.read_csv(csv_file_path)
    return df

# Function to integrate with GPT-3.5 Turbo for answer generation
def generate_answer_with_context(context, question, api_key, selected_model):
    url = "https://api.openai.com/v1/chat/completions"
    headers = {
        "Content-Type": "application/json",
        "Authorization": f"Bearer {api_key}"
    }
    prompt = f"Based on the following details: {context}, what is the concise answer for '{question}'? Keep the response brief."
    payload = {
        "model": selected_model,
        "messages": [{"role": "user", "content": prompt}],
        "temperature": 0.7
    }
    response = requests.post(url, json=payload, headers=headers)
    try:
        answer = response.json()['choices'][0]['message']['content']
        return answer.strip()
    except (KeyError, IndexError, TypeError):
        return "Error in processing the response."

# Function to perform OCR on a PDF and extract text from images
def pdf_to_images_and_extract_text(pdf_path):
    doc = fitz.open(pdf_path)
    extracted_text = ""
    for page in doc:
        pix = page.get_pixmap()
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        page_text = pytesseract.image_to_string(img)
        extracted_text += page_text + "\n"
    return extracted_text

# Function to add text to an existing PDF
def add_text_to_pdf(input_pdf_path, extracted_text, output_pdf_path):
    doc = fitz.open(input_pdf_path)
    text_blocks = extracted_text.split('\n\n')

    for page_num, page in enumerate(doc):
        text_position = fitz.Point(72, 72)
        for block in text_blocks:
            page.insert_text(text_position, block, fontsize=11)
            text_position = fitz.Point(text_position.x, text_position.y + 12)

        if page_num < len(doc) - 1:
            break

    doc.save(output_pdf_path)

# Function to extract text and their coordinates from a PDF
def extract_text_coordinates(pdf_path):
    doc = fitz.open(pdf_path)
    text_info = []

    for page_num in range(len(doc)):
        page = doc.load_page(page_num)
        text_instances = page.get_text("dict")["blocks"]

        for instance in text_instances:
            if 'lines' in instance:
                for line in instance['lines']:
                    for span in line['spans']:
                        text = span['text']
                        bbox = span['bbox']
                        text_info.append({
                            "page": page_num,
                            "text": text,
                            "bbox": bbox
                        })

    return text_info

# Function to fill form fields in a PDF with answers
def fill_pdf_form_fields(pdf_path, answers, coordinates, output_pdf_path):
    doc = fitz.open(pdf_path)
    modifications = []

    for key, value in answers.items():
        for coord in coordinates:
            if coord['text'].strip(':') == key:
                page_number = coord['page']
                x1 = coord['bbox'][2]
                y0, y1 = coord['bbox'][1], coord['bbox'][3]
                y_mid = (y0 + y1) / 2

                modifications.append({
                    'page': page_number,
                    'text': value,
                    'position': (x1 + 10, y_mid - 5),
                    'size': 10
                })
                break

    for mod in modifications:
        page = doc.load_page(mod['page'])
        text = mod['text']
        position = mod['position']
        size = mod.get('size', 11)
        page.insert_text(position, text, fontsize=size)

    doc.save(output_pdf_path)

# Example usage
pdf_path = '/Users/hope/Desktop/Final Project/24MARCH/Forms/CDC-Small-Grants-Application-Form-final-1.pdf'
csv_file_path = '/Users/hope/Desktop/Final Project/uploads/dummy_data.csv'
api_key = 'api key'
selected_model = 'gpt-3.5-turbo'

# Step 1: Extract text from PDF
extracted_text = extract_all_text(pdf_path)

# Step 2: Load CSV data
df = load_csv(csv_file_path)
data_context_dict = df.to_dict(orient='records')[0]
data_context = '. '.join([f"{key} is {value}" for key, value in data_context_dict.items()])

# Step 3: Generate answers using GPT-3.5 Turbo
keywords = [line.split(':', 1)[0] for line in extracted_text.split('\n') if ':' in line]
answers = {keyword: generate_answer_with_context(data_context_dict, keyword, api_key, selected_model) for keyword in keywords}

# Step 4: Apply OCR and extract text from images
images_text = pdf_to_images_and_extract_text(pdf_path)
extracted_text += images_text

# Step 5: Add extracted text to a new PDF
output_pdf_path_1 = '/Users/hope/Desktop/Final Project/Report Writing codes/forms/Doc1.pdf'
add_text_to_pdf(pdf_path, extracted_text, output_pdf_path_1)

# Step 6: Extract text and coordinates from the modified PDF
text_coordinates = extract_text_coordinates(output_pdf_path_1)

# Step 7: Fill form fields in the modified PDF with answers
output_pdf_path_2 = '/Users/hope/Desktop/Final Project/Report Writing codes/forms/Final_Doc.pdf'
fill_pdf_form_fields(output_pdf_path_1, answers, text_coordinates, output_pdf_path_2)

print("Process completed successfully.")
