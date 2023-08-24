import os
import re
import PyPDF2
from openpyxl import Workbook

def extract_numbers(text):
    numbers = re.findall(r'\d+\.\d+|\d+', text)
    return numbers

def process_pdf(pdf_path):
    with open(pdf_path, 'rb') as pdf_file:
        pdf_reader = PyPDF2.PdfReader(pdf_file)
        pdf_text = ''.join(page.extract_text() for page in pdf_reader.pages)
    
    numbers = extract_numbers(pdf_text)
    return numbers

def main():
    input_folder = 'input_pdfs'
    output_folder = 'output_excels'
    
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)
    
    for pdf_filename in os.listdir(input_folder):
        if pdf_filename.endswith('.pdf'):
            pdf_path = os.path.join(input_folder, pdf_filename)
            numbers = process_pdf(pdf_path)
            
            excel_filename = os.path.splitext(pdf_filename)[0] + '.xlsx'
            excel_path = os.path.join(output_folder, excel_filename)
            
            wb = Workbook()
            ws = wb.active
            ws.append(['Extracted Numbers'])
            for number in numbers:
                ws.append([number])
            wb.save(excel_path)
            print(f"Processed: {pdf_filename} -> {excel_filename}")

if __name__ == "__main__":
    main()
