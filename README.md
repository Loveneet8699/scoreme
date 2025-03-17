# Table Extraction from PDFs

## Overview
This project extracts tables from system-generated PDFs and saves the extracted data into an Excel file.

## Features
- Extracts text-based tables from PDFs
- Handles structured row alignment
- Saves extracted tables into an Excel sheet

## Prerequisites
Ensure you have Python installed. Install the required dependencies using:

```sh
pip install pdfplumber openpyxl
```

## Usage
1. Place your PDF file in the desired directory.
2. Modify `pdf_file_path` and `excel_file_path` in `Code.py` to specify input and output paths.
3. Run the script:

```sh
python Code.py
```

## Output
The extracted tables will be saved as an Excel file at the specified location.

## Notes
- Works best with system-generated PDFs containing well-structured tables.
- Does not support scanned images or highly irregular tables.


Test6.pdf file is corrupted so i can't extract the output table file for this pdf file