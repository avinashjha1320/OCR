from pdf2image import convert_from_path
import pytesseract
import re
from openpyxl import Workbook
from openpyxl.utils import get_column_letter

# Define the regex patterns for your three cases
unique_id_pattern = re.compile(r'\b[A-Z]{2,4}-[A-Za-z0-9]{4,9}[-#&*+\w]*\b')
product_description_pattern = re.compile(r'[A-Za-z0-9 ]{10,}')
product_price_pattern = re.compile(r'\b\d{3,5}\b')

# Create a list of tuples, where each tuple defines the column header and its corresponding regex pattern
# The order of the tuples in the list determines the order of the columns in the Excel sheet
column_patterns = [('Unique ID', unique_id_pattern), ('Product Description', product_description_pattern), ('Product Price', product_price_pattern)]

# Load the PDF as images
pdf_path = 'sample.pdf'
images = convert_from_path(pdf_path, dpi=200, first_page=1, last_page=3)

# Initialize the workbook and sheet
wb = Workbook()
sheet = wb.active

# Loop over the images and extract text using Pytesseract
for i, img in enumerate(images):
    text = pytesseract.image_to_string(img)

    # Initialize a list to store the row data
    row_data = []

    # Loop over the column patterns and extract the matching text
    for column_header, pattern in column_patterns:
        # Find all matches of the pattern in the text
        matches = pattern.findall(text)

        # If no matches were found, append an empty string to the row data
        if not matches:
            row_data.append('')
        # If only one match was found, append it to the row data
        elif len(matches) == 1:
            row_data.append(matches[0])
        # If multiple matches were found, concatenate them and append the result to the row data
        else:
            row_data.append(', '.join(matches).split(','))

    # Write the row data to the sheet
    max_rows = max(len(x) for x in row_data)
    for row_num in range(max_rows):
        row = []
        for column in row_data:
            try:
                row.append(column[row_num].strip())
            except IndexError:
                row.append('')
        sheet.append(row)

# Save the workbook
wb.save('output.xlsx')
