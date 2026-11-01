import os
from flask import Flask, request, send_file
from pdf2docx import Converter
import pdfplumber
import openpyxl

app = Flask(__name__)

@app.route('/')
def home():
    return "PDF Converter Server is Running Successfully!"

# 1. PDF to Word Route
@app.route('/convert-to-word', methods=['POST'])
def convert_to_word():
    if 'file' not in request.files:
        return "No file uploaded", 400
    
    file = request.files['file']
    pdf_path = "temp.pdf"
    docx_path = "output.docx"
    
    file.save(pdf_path)
    
    try:
        cv = Converter(pdf_path)
        cv.convert(docx_path, start=0, end=None)
        cv.close()
        return send_file(docx_path, as_attachment=True, download_name="converted.docx")
    except Exception as e:
        return str(e), 500
    finally:
        if os.path.exists(pdf_path): os.remove(pdf_path)

# 2. PDF to Excel Route
@app.route('/convert-to-excel', methods=['POST'])
def convert_to_excel():
    if 'file' not in request.files:
        return "No file uploaded", 400
        
    file = request.files['file']
    pdf_path = "temp_excel.pdf"
    excel_path = "output.xlsx"
    
    file.save(pdf_path)
    
    try:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Sheet1"
        
        with pdfplumber.open(pdf_path) as pdf:
            row_idx = 1
            for page in pdf.pages:
                tables = page.extract_tables()
                for table in tables:
                    for row in table:
                        col_idx = 1
                        for cell in row:
                            ws.cell(row=row_idx, column=col_idx, value=cell)
                            col_idx += 1
                        row_idx += 1
                        
        wb.save(excel_path)
        return send_file(excel_path, as_attachment=True, download_name="converted.xlsx")
    except Exception as e:
        return str(e), 500
    finally:
        if os.path.exists(pdf_path): os.remove(pdf_path)

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)