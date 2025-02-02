import re
import openpyxl
import pypdf
import os
from urllib.parse import unquote



def read_pdf(pdf_path: str) -> str:
    """
    Read data from a PDF file and store it as a single string.
    :param pdf_path: Path to the PDF file.
    :return: Extracted text from the PDF as a string.
    """

    # print(pdf_path)
    pdf_as_str = ""

    try:
        with open(pdf_path, 'rb') as pdf:
            pdf_reader = pypdf.PdfReader(pdf)  # Create a PDF reader object

            # Iterate through each page in the PDF
            for page_num in range(len(pdf_reader.pages)):
                page = pdf_reader.pages[page_num]

                # Extract text from the page
                text = page.extract_text()
                if text:
                    pdf_as_str += text

    except FileNotFoundError:
        print(f"Error: The file at {pdf_path} was not found.")
        return ""
    except Exception as e:
        print(f"An unexpected error occurred: {e}")

    return pdf_as_str


def print_from_list(data) -> None:
    """
    Read all data from iterable.
    :param data: Any iterable type.
    :return: None.
    """
    for item in data:
        print(item)


def extract_document_code(text: str, search_pattern: list = ("Document No.:", "Internal Ref. Nr.:")) -> str:
    """
    Extract the document code that is preceded by a user-defined pattern
    and ends with a newline character.

    :param text: The text in which to search for the document code.
    :param search_pattern: The list or regex pattern to search for before the document code
            (default is "Document No.:", "Internal Ref. Nr.:").
    :return: The extracted document code or None if not found.
    """

    for item in search_pattern:
        search_regex = re.escape(item) + r"([^\n]+)"

        match = re.search(search_regex, text)

        if match:
            return match.group(1).strip()

    return None


def process_excel_and_pdfs(excel_file_path: str, pdf_link_prefix: str = "", start_row: int = 2, end_row=None) -> None:
    """
    Process the Excel file and for each row, extract data from the PDF linked in the Excel.
    Write the extracted document code into the next column of the same row (column need to exist and have header).
    :param excel_file_path: Path to the Excel file.
    :param pdf_link_prefix: The prefix to be added to PDF links if they are partial.
    :param start_row: Number of the row where program starts to process, value 2 by default.
    :param end_row: Number of the row where program ends, if not specified by user program process whole file.
    :return: None
    """
    print(pdf_link_prefix)
    wb_obj = openpyxl.load_workbook(excel_file_path)
    active_sheet = wb_obj.active

    if end_row is None:
        end_row = active_sheet.max_row

    # Iterate over the rows in the Excel file
    for row in active_sheet.iter_rows(min_row=start_row, max_row=end_row, min_col=1, max_col=active_sheet.max_column):
        product_code = row[0].value  # Assuming the product code is in column A
        pdf_link = row[1].hyperlink.target if row[1].hyperlink else None  # Extract the hyperlink from column B
        if pdf_link:  # Only process if there's a valid hyperlink
            print(f"Processing PDF for Product Code: {product_code} - {pdf_link}")

            # Decode the URL-encoded characters in the file path
            decoded_pdf_link = unquote(pdf_link)

            # Join the prefix and the decoded path
            full_path = os.path.join(pdf_link_prefix, decoded_pdf_link.lstrip("\\/"))

            print(f"Decoded and full path: {full_path}")

            pdf_text = read_pdf(full_path)

            # Extract the document code from the PDF text
            document_code = extract_document_code(pdf_text)

            if document_code:
                print(f"Extracted Document Code: {document_code}")

                # Write the document code to the next column (C)
                row[2].value = document_code

    # Save the updated Excel file after processing
    wb_obj.save(f"updated_{os.path.basename(excel_file_path)}")


if __name__ == '__main__':

    pass
    # Main task

    # pdf_link_prefix = r""  # Here type PDF prefix file if needed
    # excel_file = r""  # Here type path to the excel file (Necessary)
    # process_excel_and_pdfs(excel_file, pdf_link_prefix)





