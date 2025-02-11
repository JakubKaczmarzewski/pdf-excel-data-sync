import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
import main


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
    # Read users input
    excel_path = excel_entry.get()
    pdf_prefix_path = pdf_entry.get()
    row_to_start = col_start_entry.get()
    row_to_stop = col_end_entry.get()

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


def start_process_from_folder() -> None:

    # Read users input
    folder_path = file_path_entry.get()
    excel_path = process_excel_entry.get()
    excel_path_to_save = save_excel_entry.get()

    if not folder_path and not excel_path:
        messagebox.showerror("Error", "Folder path and Excel file path are required.")
        return

    if not save_excel_entry.get():
        excel_path_to_save = excel_path
        print(f"Test {excel_path_to_save}")

    print(excel_path_to_save)

    # Setting up arguments for function

    kwargs = {"file_path": folder_path}

    kwargs["process_excel_path"] = excel_path
    if excel_path_to_save:
        kwargs["excel_save_path"] = excel_path_to_save

    try:
        main.process_folder_write_pdfs_data_to_excel(**kwargs)
        messagebox.showinfo("Success", "Processing completed successfully.")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")


main_window = tk.Tk()
main_window.title("Process Excel and PDF")
main_window.geometry('700x550')

# Main window configure:
main_window.rowconfigure(0, weight=1)
main_window.rowconfigure(1, weight=1)
main_window.rowconfigure(2, weight=1)

main_window.columnconfigure(0, weight=1)
main_window.columnconfigure(1, weight=1)

# Frame: Process Excel Section
excel_frame = ttk.LabelFrame(main_window, text="Process Excel and write Document codes", padding=(10, 10))
excel_frame.grid(row=0, column=0, padx=10, pady=10, sticky="ew")

excel_label = ttk.Label(excel_frame, text="Excel File Path:")
excel_label.grid(row=0, column=0, sticky="w", pady=5)
excel_entry = ttk.Entry(excel_frame, width=50)
excel_entry.grid(row=0, column=1, pady=5)

pdf_label = ttk.Label(excel_frame, text="PDF Prefix Path:")
pdf_label.grid(row=1, column=0, sticky="w", pady=5)
pdf_entry = ttk.Entry(excel_frame, width=50)
pdf_entry.grid(row=1, column=1, pady=5)

col_start_label = ttk.Label(excel_frame, text="Starting Column:")
col_start_label.grid(row=2, column=0, sticky="w", pady=5)
col_start_entry = ttk.Entry(excel_frame, width=50)
col_start_entry.grid(row=2, column=1, pady=5)

col_end_label = ttk.Label(excel_frame, text="Ending Column:")
col_end_label.grid(row=3, column=0, sticky="w", pady=5)
col_end_entry = ttk.Entry(excel_frame, width=50)
col_end_entry.grid(row=3, column=1, pady=5)

run_button = ttk.Button(excel_frame, text="Start Processing", command=start_process)
run_button.grid(row=4, column=0, columnspan=2, pady=10, sticky='w')

# Frame: Process File Section
file_frame = ttk.LabelFrame(main_window, text="Process from file and write all data to Excel", padding=(10, 10))
file_frame.grid(row=1, column=0, padx=10, pady=10, sticky="ew")

file_path_label = ttk.Label(file_frame, text="File Path:")
file_path_label.grid(row=0, column=0, sticky="w", pady=5)
file_path_entry = ttk.Entry(file_frame, width=50)
file_path_entry.grid(row=0, column=1, pady=5)

process_excel_label = ttk.Label(file_frame, text="Path to Excel to \n process:")
process_excel_label.grid(row=1, column=0, sticky="w", pady=5)
process_excel_entry = ttk.Entry(file_frame, width=50)
process_excel_entry.grid(row=1, column=1, pady=5)

save_excel_label = ttk.Label(file_frame, text="Path to save\n Excel changes:")
save_excel_label.grid(row=2, column=0, sticky="w", pady=5)
save_excel_entry = ttk.Entry(file_frame, width=50)
save_excel_entry.grid(row=2, column=1, pady=5)

file_run_button = ttk.Button(file_frame, text="Start Processing", command=start_process_from_folder)
file_run_button.grid(row=3, column=0, columnspan=2, pady=10, sticky='w')

# Instruction frame
instruction_frame = ttk.LabelFrame(main_window, text="Instruction", padding=(10, 10))
instruction_frame.grid(row=2, column=0, padx=10, pady=10, sticky="ew")

instruction_button = ttk.Button(instruction_frame, text="How to Use", command=show_instruction)
instruction_button.grid(row=0, column=0)

# Saving option frame
saving_option_frame = ttk.LabelFrame(main_window, text="Saving file options")
saving_option_frame.grid(row=0, column=1, sticky='n', pady=10)

rb_value = tk.IntVar()
rb_value.set(2)

# Radio save buttons
overwrite_save_rb = ttk.Radiobutton(saving_option_frame, text='Overwrite file', value=1, variable=rb_value)
overwrite_save_rb.grid(row=0, column=0, padx=5, pady=10)
overwrite_save_rb = ttk.Radiobutton(saving_option_frame, text='Create updated copy', value=2, variable=rb_value)
overwrite_save_rb.grid(row=0, column=1, padx=5, pady=10)

if rb_value.get() == 2:
    print("Value 2 ")
elif rb_value.get() == 1:
    print("Value 1")
main_window.mainloop()
