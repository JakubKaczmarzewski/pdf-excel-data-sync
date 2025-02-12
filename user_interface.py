import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
from tkinter import scrolledtext
import main


def show_instruction() -> None:
    """
    Print instruction text on user's screen.
    :return: None
    """
    instruction_text = """
    User Manual for the Program:
    
    1) Process PDF Files and Write Data to Excel:
    --------------------------------------------------
    **About this functionality:**
    This feature processes all PDF files from a specified folder. For each file, it extracts **Product Codes** and a **Document Number**, then moves the file to a user-defined location or a default "After Processed" folder. Each processed file is logged in the Excel sheet with its Hyperlink and Document Number alongside corresponding Product Codes.
    
    **Steps to use this functionality:**
    1.1 In the **`File Path`** field, paste the absolute path to the folder containing the PDF files you want to process.
        Example:
        `C:\\Users\\User123\\Desktop\Folder_with_my_pdfs`
    
    1.2 In the **`Path to Excel to process`** field, provide the absolute path to the Excel file to be updated.
        Example:
        `C:\\Users\\User123\\Desktop\excel_file.xlsx`
    
    1.3 In the **`Path to save Excel changes`** field, specify the folder where you want the modified Excel file to be saved.
        - If left empty, the changes will be saved in the same directory as the original Excel file.
    
    1.4 In the **`Path to save processed PDF's`** field, specify the absolute path to the folder where successfully processed files should be stored.
        Example:
        `C:\\Users\\User123\\Desktop\\Processed_PDFs`
        **Note:** If you want to use the default location, leave this field empty.
    
    1.5 In the **`Excel saving file options`**, select one of the following:
        - **Overwrite a file:** Saves changes to the same Excel file, replacing the previous version. **Warning:** This will overwrite the original file.
        - **Create updated copy (default):** Saves the updated Excel file with the prefix `updated_` in the same location as the original.
          Example:
          Original file: `C:\\Users\\User123\\Desktop\\excel_file.xlsx`
          Updated file: `C:\\Users\\User123\\Desktop\\updated_excel_file.xlsx`
    
    1.6 Click the **`Start Processing`** button to run the program. The processing time will vary depending on the number of files. Once completed, a notification window will display the program's status.
    
    --------------------------------------------------
    
    2) Process Excel and Write Document Codes:
    --------------------------------------------------
    **About this functionality:**
    This feature reads data from an Excel file, extracts text from related PDF files, and writes the results back into the Excel file.
    
    **Steps to use this functionality:**
    2.1 In the **`Path to Excel to process`** field, enter the full path to the Excel file you want to update.
    
    2.2 If the Excel file contains hyperlinks to PDF files, provide the base path (prefix) to the folder containing the PDF files in the **`PDF File Path Prefix`** field.
    
    2.3 In the **`Starting Row`** field, enter the row number where processing should begin.
        - Leave this empty to start from the second row (default).
    
    2.4 In the **`Stopping Column`** field, enter the column number where processing should end.
        - Leave this empty to process the entire file.
    
    2.5 Click the **`Start`** button to begin processing. Once finished, the program will display a status message.
    
    --------------------------------------------------
    
    **Tips for Success:**
    - Always provide absolute paths to avoid errors.
    - Ensure the Excel file is closed before starting the program.
    - Use the "Create updated copy" option to preserve your original files.
    
    Thank you for using my program!
    """
    # Create a new top-level window
    manual_window = tk.Toplevel()
    manual_window.title("User Manual")

    # Set minimum size and center it on the screen
    manual_window.geometry("600x400")
    manual_window.resizable(True, True)

    # Create a scrolled text widget
    text_area = scrolledtext.ScrolledText(manual_window, wrap=tk.WORD, font=("Arial", 10))
    text_area.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

    # Insert the manual text
    text_area.insert(tk.END, instruction_text)

    # Make the text area read-only
    text_area.config(state=tk.DISABLED)

    # Add a close button
    close_button = tk.Button(manual_window, text="Close", command=manual_window.destroy)
    close_button.pack(pady=10)


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

    # Error catching
    if not excel_path:
        messagebox.showerror("Error", "Excel file path is required.")
        return
    # if row_to_start > row_to_stop:
    #     messagebox.showerror("Error", "The starting row value must be less than the stop value.")
    #     return

    # Przygotowanie argumentów dla funkcji
    kwargs = {"excel_file_path": excel_path}

    if pdf_prefix_path:
        kwargs["pdf_link_prefix"] = pdf_prefix_path
    if row_to_start:
        kwargs["start_row"] = int(row_to_start)
    if row_to_stop:
        kwargs["end_row"] = int(row_to_stop)

    save_option = rb_value.get()
    kwargs["save_option"] = save_option

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
    pdfs_folder_path_to_save = save_file_entry.get()

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
    save_option = rb_value.get()
    kwargs["save_option"] = save_option
    if save_file_entry:
        kwargs["move_to_folder"] = pdfs_folder_path_to_save

    try:
        main.process_folder_write_pdfs_data_to_excel(**kwargs)
        messagebox.showinfo("Success", "Processing completed successfully.")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")


main_window = tk.Tk()
main_window.title("Process Excel and PDF v1.0")
main_window.geometry('700x550')

# Main window configure:
main_window.rowconfigure(0, weight=1)
main_window.rowconfigure(1, weight=1)
main_window.rowconfigure(2, weight=1)

main_window.columnconfigure(0, weight=1)
main_window.columnconfigure(1, weight=1)

# Frame: Process Excel Section
excel_frame = ttk.LabelFrame(main_window, text="Process Excel and write Document codes", padding=(10, 10))
excel_frame.grid(row=1, column=0, padx=10, pady=10, sticky="ew")

excel_label = ttk.Label(excel_frame, text="Excel File Path:")
excel_label.grid(row=0, column=0, sticky="w", pady=5)
excel_entry = ttk.Entry(excel_frame, width=50)
excel_entry.grid(row=0, column=1, pady=5)

pdf_label = ttk.Label(excel_frame, text="PDF Prefix Path:")
pdf_label.grid(row=1, column=0, sticky="w", pady=5)
pdf_entry = ttk.Entry(excel_frame, width=50)
pdf_entry.grid(row=1, column=1, pady=5)

col_start_label = ttk.Label(excel_frame, text="Starting Row:")
col_start_label.grid(row=2, column=0, sticky="w", pady=5)
col_start_entry = ttk.Entry(excel_frame, width=50)
col_start_entry.grid(row=2, column=1, pady=5)

col_end_label = ttk.Label(excel_frame, text="Ending Row:")
col_end_label.grid(row=3, column=0, sticky="w", pady=5)
col_end_entry = ttk.Entry(excel_frame, width=50)
col_end_entry.grid(row=3, column=1, pady=5)

run_button = ttk.Button(excel_frame, text="Start Processing", command=start_process)
run_button.grid(row=4, column=0, columnspan=2, pady=10, sticky='w')

# Frame: Process File Section
file_frame = ttk.LabelFrame(main_window, text="Process from file and write all data to Excel", padding=(10, 10))
file_frame.grid(row=0, column=0, padx=10, pady=10, sticky="ew")

file_path_label = ttk.Label(file_frame, text="File Path:")
file_path_label.grid(row=0, column=0, sticky="w", pady=5)
file_path_entry = ttk.Entry(file_frame, width=50)
file_path_entry.grid(row=0, column=1, pady=5)

process_excel_label = ttk.Label(file_frame, text="Path to Excel to \nprocess:")
process_excel_label.grid(row=1, column=0, sticky="w", pady=5)
process_excel_entry = ttk.Entry(file_frame, width=50)
process_excel_entry.grid(row=1, column=1, pady=5)

save_excel_label = ttk.Label(file_frame, text="Path to save\nExcel changes:")
save_excel_label.grid(row=2, column=0, sticky="w", pady=5)
save_excel_entry = ttk.Entry(file_frame, width=50)
save_excel_entry.grid(row=2, column=1, pady=5)

save_file_label = ttk.Label(file_frame, text="Path to save\nprocessed PDF's.")
save_file_label.grid(row=3, column=0, sticky='w', pady=5)
save_file_entry = ttk.Entry(file_frame, width=50)
save_file_entry.grid(row=3, column=1, columnspan=2, pady=10, sticky='w')

file_run_button = ttk.Button(file_frame, text="Start Processing", command=start_process_from_folder)
file_run_button.grid(row=4, column=0, columnspan=2, pady=10, sticky='w')

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

main_window.mainloop()
