from flask import Flask, request, render_template, send_file
from werkzeug.utils import secure_filename
from pdf2image import convert_from_path
import pytesseract
from docx import Document
from PIL import Image
import os

app = Flask(__name__)
UPLOAD_FOLDER = 'uploads'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

# Set the path to the Tesseract executable
pytesseract.pytesseract.tesseract_cmd = 'C:\\Program Files\\Tesseract-OCR\\tesseract.exe'

def image_to_text(image_path):
    image = Image.open(image_path)
    text = pytesseract.image_to_string(image, lang='ben')
    return text

def pdf_to_text(pdf_path):
    images = convert_from_path(pdf_path)
    text = ''
    for image in images:
        text += pytesseract.image_to_string(image, lang='ben')
    return text

def save_text_to_file(text, file_path):
    with open(file_path, 'w', encoding='utf-8') as file:
        file.write(text)

def save_text_to_doc(text, doc_path):
    doc = Document()
    doc.add_paragraph(text)
    doc.save(doc_path)

@app.route('/', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        if 'file' not in request.files:
            return 'No file part'
        file = request.files['file']
        if file.filename == '':
            return 'No selected file'
        if file:
            filename = secure_filename(file.filename)
            file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            file.save(file_path)

            if file_path.lower().endswith('.pdf'):
                extracted_text = pdf_to_text(file_path)
            else:
                extracted_text = image_to_text(file_path)

            return render_template('result.html', extracted_text=extracted_text)
    return render_template('index.html')

@app.route('/download', methods=['POST'])
def download_file():
    text = request.form['extracted_text']
    download_type = request.form['download_type']
    if download_type == 'txt':
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], 'output.txt')
        save_text_to_file(text, file_path)
        return send_file(file_path, as_attachment=True)
    elif download_type == 'docx':
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], 'output.docx')
        save_text_to_doc(text, file_path)
        return send_file(file_path, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)
