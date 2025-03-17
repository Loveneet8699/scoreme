import pdfplumber
import openpyxl

def extract_pdf_tables(pdf_path):
    extracted_data = []
    with pdfplumber.open(pdf_path) as pdf:
        for page_idx, page in enumerate(pdf.pages, start=1):
            page_words = page.extract_words()
            if page_words:
                structured_rows = {}
                for word in page_words:
                    row_position = round(word['top'])
                    if row_position not in structured_rows:
                        structured_rows[row_position] = []
                    structured_rows[row_position].append(word['text'])
                sorted_table = [structured_rows[key] for key in sorted(structured_rows.keys())]
                extracted_data.extend(sorted_table)
            
    return extracted_data

def save_tables_to_excel(table_data, output_path):
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    
    if table_data:
        for row_idx, row in enumerate(table_data, start=1):
            for col_idx, cell in enumerate(row, start=1):
                sheet.cell(row=row_idx, column=col_idx, value=cell)
        workbook.save(output_path)

pdf_file_path = r"D:\Scoreme\test3.pdf"
excel_file_path = r"D:\Scoreme\output_tables.xlsx"

tables = extract_pdf_tables(pdf_file_path)
save_tables_to_excel(tables, excel_file_path)
