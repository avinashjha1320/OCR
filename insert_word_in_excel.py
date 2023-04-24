from pdf2image import convert_from_path
import pytesseract
from PIL import Image
from openpyxl import Workbook

# Load PDF and convert to image
pdf_path = 'sample.pdf'
images = convert_from_path(pdf_path, dpi=200, first_page=1, last_page=3)

# Create workbook and worksheet
wb = Workbook()
ws = wb.active
row = 1

# Loop through images and extract words
for i, image in enumerate(images):
    image_path = f'page_{i+1}.jpg'
    image.save(image_path, 'JPEG')

    # Extract text from image using OCR
    text = pytesseract.image_to_string(Image.open(image_path))

    # Split text into words
    words = text.split()

    # Write each word to worksheet
    for word in words:
        ws.cell(row=row, column=1, value=word)
        row += 1

# Save workbook
wb.save('words.xlsx')