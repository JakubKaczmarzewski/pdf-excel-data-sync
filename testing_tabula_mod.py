import tabula
import pandas as pd

pdf_path = r"C:\Users\670336256\Desktop\2021A0117.pdf"

dfs = tabula.read_pdf(pdf_path, pages='all', multiple_tables=True)
print(type(dfs[0]))


