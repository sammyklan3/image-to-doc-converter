import os
import glob
import pytesseract
from PIL import Image
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT

# Path to the folder containing the PDF images
pdf_images_folder = "images"

pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

# Path to the output Word document
output_docx = "output.docx"

# Function to extract text from image using OCR
def extract_text_from_image(image_path):
    try:
        return pytesseract.image_to_string(Image.open(image_path))
    except Exception as e:
        print(f"Error occurred while processing {image_path}: {str(e)}")
        return ""

# Function to create a Word document and format the text
def create_word_document(texts):
    doc = Document()
    for text in texts:
        paragraph = doc.add_paragraph()
        run = paragraph.add_run(text)
        font = run.font
        font.size = Pt(12)  # Change the font size as needed
        paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT  # Align text to left
    try:
        doc.save(output_docx)
        print("Word document created successfully.")
    except Exception as e:
        print(f"Error occurred while saving the Word document: {str(e)}")

# Main function
def main():
    texts = []
    image_files = glob.glob(os.path.join(pdf_images_folder, "*.jpg")) + glob.glob(os.path.join(pdf_images_folder, "*.png"))
    if not image_files:
        print("No image files found in the specified folder.")
        return
    for image_path in image_files:
        text = extract_text_from_image(image_path)
        texts.append(text)
    create_word_document(texts)

if __name__ == "__main__":
    main()
