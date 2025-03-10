import os
import pandas as pd
import sys
from PyPDF2 import PdfReader

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
                # Procurar o valor desejado
                value = extract_text_between_keywords(page_text, [("(=) Valor do\nDocumento:", "(-) Desconto /"), ("ValorR$ ", "Desconto:"), ("Valor (R$): ", "Finalidade:"), ("Valor (R$):", "Uso")])
                if value is None:  # Se o valor não for encontrado, tente outra combinação de palavras-chave
                    value = extract_text_between_keywords(page_text, [("Valor (R$):","Uso")])
                data.append({"Nome do Arquivo": filename, "Valor": value})

    return data

def save_to_excel(data, output_excel_path):
    df = pd.DataFrame(data)
    df.to_excel(output_excel_path, index=False)
    print(f"Dados salvos em: {output_excel_path}")

if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("Por favor, forneça os seguintes parâmetros: caminho da pasta de entrada e caminho da pasta de saída.")
        sys.exit(1)

    input_folder = sys.argv[1]  # Caminho da pasta de entrada dos PDFs
    output_folder = sys.argv[2]  # Caminho da pasta de saída para o Excel

    if not os.path.exists(input_folder):
        print(f"A pasta de entrada {input_folder} não existe.")
        sys.exit(1)

    if not os.path.exists(output_folder):
        print(f"A pasta de saída {output_folder} não existe.")
        sys.exit(1)

    # Busca os valores nos PDFs
    data = search_value_in_pdfs(input_folder)

    # Caminho para o arquivo Excel de saída
    output_excel_path = os.path.join(output_folder, "output_data.xlsx")

    # Salva os dados em um arquivo Excel
    save_to_excel(data, output_excel_path)
