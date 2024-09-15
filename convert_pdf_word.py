from pdf2docx import Converter

def convert_pdf_to_word(pdf_path, output_path):
    # Initialize the converter
    converter = Converter(pdf_path)

    # Convert the PDF to Word
    converter.convert(output_path, start=0, end=None)

    # Close the converter
    converter.close()

if __name__ == "__main__":
    # Replace 'input.pdf' with the path to your PDF file
    input_pdf_path = 'Data/BCU-application-form.pdf'

    # Replace 'output.docx' with the desired output Word file path
    output_word_path = 'Data/BCU-application-form.docx'

    convert_pdf_to_word(input_pdf_path, output_word_path)
