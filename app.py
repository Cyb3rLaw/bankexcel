from flask import Flask, render_template, request, redirect, url_for, send_file
import os
import pandas as pd
from pdf2xlsx import pdf2xlsx
from PIL import Image
import pytesseract
import spacy
import datefinder
import re

app = Flask(__name__)
# Other codes
if __name__ == '__main__':
    app.run()

UPLOAD_FOLDER = 'uploads'
OUTPUT_FOLDER = 'outputs'

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['OUTPUT_FOLDER'] = OUTPUT_FOLDER

# Load SpaCy English model for named entity recognition
nlp = spacy.load("en_core_web_sm")

# Define regular expressions for pattern matching
date_pattern = re.compile(r'\d{1,2}/\d{1,2}/\d{2,4}')
debit_credit_pattern = re.compile(r'\d+\.\d{2}')
balance_pattern = re.compile(r'Balance:?\s*\d+\.\d{2}')


def ocr(pdf_path):
    # Convert PDF to images
    images = pdf2images(pdf_path)

    # Perform OCR on each image
    text_content = ''
    for img in images:
        text_content += pytesseract.image_to_string(img, lang='eng')

    return text_content


def pdf2images(pdf_path):
    images = []
    pdf_images = pdf2xlsx.extract_images(pdf_path)

    for i, pdf_image in enumerate(pdf_images):
        img = Image.open(pdf_image)
        images.append(img)
        img.save(os.path.join(app.config['UPLOAD_FOLDER'], f'page_{i + 1}.png'))

    return images


def extract_entities(text):
    doc = nlp(text)
    entities = {"dates": [], "other": []}

    for ent in doc.ents:
        if ent.label_ == "DATE":
            entities["dates"].append(ent.text)
        else:
            entities["other"].append((ent.text, ent.label_))

    return entities


def extract_bank_statement_data(text):
    # Extract fields using regular expressions
    dates = date_pattern.findall(text)
    debit_credit_matches = debit_credit_pattern.findall(text)
    balance_matches = balance_pattern.findall(text)

    data = []
    for date in dates:
        data_point = {"Date": date, "Particulars": "", "Debit": "", "Credit": "", "Running Balance": ""}
        
        for match in debit_credit_matches:
            if match in text:
                index = text.find(match)
                preceding_text = text[:index].strip().rsplit('\n', 1)[-1]
                if "debit" in preceding_text.lower():
                    data_point["Debit"] = match
                elif "credit" in preceding_text.lower():
                    data_point["Credit"] = match

        for match in balance_matches:
            if match in text:
                index = text.find(match)
                preceding_text = text[:index].strip().rsplit('\n', 1)[-1]
                if "balance" in preceding_text.lower():
                    data_point["Running Balance"] = match

        data.append(data_point)

    return data


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/upload', methods=['POST'])
def upload():
    if 'file' not in request.files:
        return redirect(request.url)

    file = request.files['file']

    if file.filename == '':
        return redirect(request.url)

    if file:
        pdf_path = os.path.join(app.config['UPLOAD_FOLDER'], 'input.pdf')
        file.save(pdf_path)

        # OCR text extraction
        ocr_text = ocr(pdf_path)

        # Save OCR text to a text file
        ocr_text_path = os.path.join(app.config['OUTPUT_FOLDER'], 'ocr_text.txt')
        with open(ocr_text_path, 'w') as f:
            f.write(ocr_text)

        # Extract named entities using SpaCy
        entities = extract_entities(ocr_text)

        # Extract bank statement data based on user's choice
        extract_choice = request.form.get('extract_choice', 'all')
        if extract_choice == 'all':
            extracted_data = extract_entities(ocr_text)
        else:
            extracted_data = extract_bank_statement_data(ocr_text)

        # Convert extracted data to Excel using pandas
        excel_path = os.path.join(app.config['OUTPUT_FOLDER'], 'output.xlsx')
        df = pd.DataFrame(extracted_data)
        df.to_excel(excel_path, index=False)

        # Delete the uploaded PDF after conversion
        os.remove(pdf_path)

        return render_template('preview.html', ocr_text=ocr_text, entities=entities, extracted_data=extracted_data, excel_path=excel_path)


@app.route('/download_excel')
def download_excel():
    return render_template('download.html')


@app.route('/download')
def download():
    excel_path = os.path.join(app.config['OUTPUT_FOLDER'], 'output.xlsx')
    return send_file(excel_path, as_attachment=True)


if __name__ == '__main__':
    if not os.path.exists(UPLOAD_FOLDER):
        os.makedirs(UPLOAD_FOLDER)
    if not os.path.exists(OUTPUT_FOLDER):
        os.makedirs(OUTPUT_FOLDER)

    app.run(debug=True)
