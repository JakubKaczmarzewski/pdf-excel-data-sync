# pdf-excel-data-sync

**PdfExcelDataSync** is an application that automatically processes data from an Excel file, utilizing linked PDF files. The main task of the application is to extract specific information from PDF files, such as document codes, and save this data into the corresponding cells in the Excel file. The application is designed to work in a Windows environment, using Python libraries such as `openpyxl` (for handling Excel files) and `pypdf` (for processing PDF files).

## Key Features:
1. **PDF Reading**: The `read_pdf` function allows you to read the content of a PDF file and store it as a string. It uses the `pypdf` library to extract text from each page of the PDF.
   
2. **Document Code Extraction**: The `extract_document_code` function uses regular expressions to extract the document code from the PDF text based on a defined pattern (e.g., "Document No." or "Internal Ref. Nr.").

3. **Excel File Processing**: The `process_excel_and_pdfs` function processes the Excel file by linking each row to the corresponding PDF file based on the hyperlink in a column. After reading the PDF content, it saves the extracted document code into the next column of the same row.

4. **Error Handling**: The program handles cases where a PDF file is missing or other unexpected errors occur during the file reading process.

## How It Works:
1. The script reads data from an Excel file that contains links to PDF files (saved in one of the columns).
2. For each link, the script opens the corresponding PDF file, extracts the text, and retrieves the document code.
3. The extracted document code is saved in the corresponding cell of the Excel file.
4. After processing all rows, the updated Excel file is saved with the new data.

## Requirements:
- Python 3.x
- Libraries: `openpyxl`, `pypdf`, `re`

## Usage Example:
The application is designed to work with an Excel file located on a network server or locally. After running the script, the user provides the path to the Excel file and optionally the prefix path to the PDF files. The application automatically processes all rows in the Excel file, extracts data from the PDF files, and saves it back to the Excel file.


