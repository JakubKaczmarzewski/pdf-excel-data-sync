pdf_link_prefix = r"\\lux.intra.lighting.com\PL-PIL001\TWWT-Box1\Obsługa celna\CE"
print(pdf_link_prefix)
pdf_link = "Dokumenty\\dok.pdf"
# excel_file = r"X:\TWWT-Box1\Obsługa celna\CE\Lista kodów.xlsx"  # plik docelowy (z dysku)
# print(excel_file)


pdf_link = pdf_link_prefix + "\\" + pdf_link.lstrip("\\")
print(pdf_link)

