from PIL import Image
import pytesseract
import pdf2image

pdf_path = 'sample.pdf'
images = pdf2image.convert_from_path(pdf_path, dpi=200)

for i, image in enumerate(images):
    image_path = f'page_{i}.jpg'
    image.save(image_path, 'JPEG')

    text = pytesseract.image_to_string(image)
    print(f'Page {i+1} text:')
    print(text)