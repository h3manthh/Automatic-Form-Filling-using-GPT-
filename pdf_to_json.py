import fitz
import json

def pdf_to_json(pdf_path, json_path):
    doc = fitz.open(pdf_path)
    data = {}

    for page_num in range(doc.page_count):
        page = doc[page_num]
        text = page.get_text()

        lines = text.split('\n')
        for line in lines:
            if ':' in line:
                key, value = map(str.strip, line.split(':', 1))
                data[key] = value

    with open(json_path, 'w') as json_file:
        json.dump(data, json_file, indent=2)

    print(f"PDF converted to JSON. Output saved to {json_path}")

if __name__ == "__main__":
    pdf_path = "/Users/hope/Desktop/Final Project/Forms/CDC-Small-Grants-Application-Form-final-1.pdf"  # Replace with your PDF file path
    json_path = "/Users/hope/Desktop/Final Project/Output/March Output/output.json"  # Replace with the desired JSON output file path

    pdf_to_json(pdf_path, json_path)
