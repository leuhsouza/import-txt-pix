import os
import pandas as pd
from flask import Flask, render_template, request, send_from_directory
from PyPDF2 import PdfReader
import tempfile
import shutil

app = Flask(__name__)

# Definir diretório temporário para arquivos carregados
UPLOAD_FOLDER = tempfile.mkdtemp()  # Pasta temporária para armazenar os arquivos enviados
OUTPUT_FOLDER = tempfile.mkdtemp()  # Pasta temporária para armazenar o arquivo de saída (Excel)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

def extract_text_between_keywords(text, keyword_pairs):
    for start_keyword, end_keyword in keyword_pairs:
        start_index = text.find(start_keyword)
        if start_index != -1:
            text = text[start_index + len(start_keyword):]
            end_index = text.find(end_keyword)
            if end_index != -1:
                return text[:end_index].strip()
    return None

def search_value_in_pdfs(input_folder):
    data = []

    for filename in os.listdir(input_folder):
        if filename.endswith(".pdf"):
            pdf_path = os.path.join(input_folder, filename)
            with open(pdf_path, "rb") as input_pdf:
                pdf_reader = PdfReader(input_pdf)
                page_text = ""
                for page in pdf_reader.pages:
                    page_text += page.extract_text()
                value = extract_text_between_keywords(page_text, [("(=) Valor do\nDocumento:", "(-) Desconto /"), ("ValorR$ ", "Desconto:"), ("Valor (R$): ", "Finalidade:"), ("Valor (R$):", "Uso")])
                if value is None:
                    value = extract_text_between_keywords(page_text, [("Valor (R$):","Uso")])
                data.append({"Nome do Arquivo": filename, "Valor": value})
    return data

def save_to_excel(data, output_excel_path):
    df = pd.DataFrame(data)
    df.to_excel(output_excel_path, index=False)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'pdf_file' not in request.files:
        return "Nenhum arquivo PDF foi enviado.", 400
    
    pdf_file = request.files['pdf_file']
    if pdf_file.filename == '':
        return "Nenhum arquivo selecionado.", 400
    
    # Salva o arquivo PDF enviado
    input_folder = os.path.join(UPLOAD_FOLDER, 'input')
    if not os.path.exists(input_folder):
        os.makedirs(input_folder)

    file_path = os.path.join(input_folder, pdf_file.filename)
    pdf_file.save(file_path)

    # Processa o PDF
    data = search_value_in_pdfs(input_folder)
    
    # Cria arquivo Excel de saída
    output_excel_path = os.path.join(OUTPUT_FOLDER, "output_data.xlsx")
    save_to_excel(data, output_excel_path)

    # Retorna o link para download do arquivo gerado
    return send_from_directory(directory=OUTPUT_FOLDER, path="output_data.xlsx", as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0')
