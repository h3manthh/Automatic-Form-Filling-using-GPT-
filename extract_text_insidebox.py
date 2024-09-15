import fitz
import pytesseract
from PIL import Image
import cv2
import numpy as np

def locate_rectangular_boxes_with_ocr(pdf_path):
    # Open the PDF file
    doc = fitz.open(pdf_path)

    all_rectangular_boxes = []  # Initialize a list to store coordinates of potential rectangular boxes

    # Iterate through each page in the PDF
    for page_number in range(doc.page_count):
        # Convert the PDF page to an image
        page = doc[page_number]
        pix = page.get_pixmap()
        image = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)

        # Perform OCR on the image
        ocr_text = pytesseract.image_to_string(image)

        # Convert the image to a NumPy array for OpenCV processing
        img_np = np.array(image)

        # Convert to grayscale
        gray = cv2.cvtColor(img_np, cv2.COLOR_BGR2GRAY)

        # Apply adaptive thresholding to create a binary image
        binary = cv2.adaptiveThreshold(gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY_INV, 11, 4)

        # Find contours in the binary image
        contours, _ = cv2.findContours(binary, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)

        # Iterate through contours to find rectangular boxes
        for contour in contours:
            # Approximate the contour to a polygon
            epsilon = 0.04 * cv2.arcLength(contour, True)
            approx = cv2.approxPolyDP(contour, epsilon, True)

            # Filter based on the number of vertices (4 for rectangles)
            if len(approx) == 4:
                x, y, w, h = cv2.boundingRect(contour)

                # Filter based on area (you may need to adjust this threshold)
                if 1000 < cv2.contourArea(contour) < 50000:
                    all_rectangular_boxes.append((page_number + 1, (x, y, w, h)))

    print(f"Rectangular boxes found on each page: {all_rectangular_boxes}")
    return all_rectangular_boxes

def extract_text_within_boxes(pdf_path, rectangular_boxes):
    # Open the PDF file
    doc = fitz.open(pdf_path)

    all_extracted_text = []  # Initialize a list to store text within identified boxes

    # Iterate through each rectangular box and extract text using PyMuPDF
    for page_number, box_coords in rectangular_boxes:
        page = doc[page_number - 1]  # PyMuPDF page indices start from 0
        page_text = page.get_text("text", clip=fitz.Rect(*box_coords))

        all_extracted_text.append(page_text)

    print(f"Text extracted within identified rectangular boxes: {all_extracted_text}")
    return all_extracted_text

# Specify the path to your PDF file
pdf_path = '/Users/hope/Desktop/Final Project/Data/BCU-application-form.pdf'

# Locate rectangular boxes using OCR and image processing
rectangular_boxes = locate_rectangular_boxes_with_ocr(pdf_path)

# Extract text within identified rectangular boxes using PyMuPDF
extracted_text_within_boxes = extract_text_within_boxes(pdf_path, rectangular_boxes)

# Print the extracted text within identified rectangular boxes
for idx, text in enumerate(extracted_text_within_boxes, start=1):
    print(f"Text within Rectangular Box {idx}:\n{text}")
