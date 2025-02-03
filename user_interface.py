import tkinter as tk
from tkinter import messagebox
import main


# Cele interface'u:
# 1 Możliwość wpisywania ścieżki do pliku excela
# 2 Możliwość wpisywania przedrostku do pliku PDF
# 3 Możliwość ustawienia startu oraz limitu przetwarzanych column w excelu
                                        # (zmiana w pliku main)
# 4 Wyświetlanie instrukcji po kliknieciu w instrukcje
# 5 Przycisk do uruchomienia całej aplikacji

def show_instruction() -> None:
    """
    Print instruction text on user's screen.
    :return: None
    """
    instruction_text = """
    Program User Manual:

    1.Enter the full path to the Excel file you want to process.
    2.If Excel file contain hyperlink enter prefix of full path to the directory containing the PDF files.
    3. Enter the number of the row you want to start from. If you want to start from the second row leave this empty.
    4. Enter the number of the column where you want to stop processing. If you want to process whole file, 
    leave this empty.
    5.Click 'Start' to begin processing the files.
    The program reads data from the Excel file, extract text from PDF combines it
    and saves the results in a new Excel file.
    """
    messagebox.showinfo("How to use this app", instruction_text)


def start_process() -> None:
    """
    Read user's input from the tkinter screen and run whole app.

    :return: None
    """
    # Pobieranie danych z pól tekstowych
    excel_path = excel_entry.get()
    pdf_prefix_path = pdf_entry.get()
    row_to_start = col_start_entry.get()
    row_to_stop = col_end_entry.get()

    # Sprawdź, czy podano niezbędne dane
    if not excel_path:
        messagebox.showerror("Error", "Excel file path is required.")
        return

    # Przygotowanie argumentów dla funkcji
    kwargs = {"excel_file_path": excel_path}

    if pdf_prefix_path:
        kwargs["pdf_link_prefix"] = pdf_prefix_path
    if row_to_start:
        kwargs["start_row"] = int(row_to_start)
    if row_to_stop:
        kwargs["end_row"] = int(row_to_stop)

    # Wywołanie funkcji z dynamicznie utworzonymi argumentami
    try:
        main.process_excel_and_pdfs(**kwargs)
        messagebox.showinfo("Success", "Processing completed successfully.")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")



main_window = tk.Tk()

main_window.title("Process Excel and Pdf")
main_window.geometry('640x480-8-200')

instruction_button = tk.Button(main_window, text="?", command=show_instruction)
instruction_button.grid(row=0, column=3, sticky=tk.E)

excel_label = tk.Label(main_window, text="1.Excel File path")
excel_label.grid(row=1, column=0)
excel_entry = tk.Entry(main_window, width=40)
excel_entry.grid(row=1, column=1)

pdf_label = tk.Label(main_window, text="2.PDF prefix path")
pdf_label.grid(row=2, column=0)
pdf_entry = tk.Entry(main_window, width=40)
pdf_entry.grid(row=2, column=1)

col_start = tk.Label(main_window, text="3.Starting column")
col_start.grid(row=3, column=0)

col_start_entry = tk.Entry(main_window, width=40)
col_start_entry.grid(row=3, column=1)

col_end = tk.Label(main_window, text="4.Ending column")
col_end.grid(row=4, column=0)
col_end_entry = tk.Entry(main_window, width=40)
col_end_entry.grid(row=4, column=1)

run_button = tk.Button(main_window, text="Start processing", command=start_process)
run_button.grid(row=5, column=3, sticky=tk.S)

main_window.mainloop()
