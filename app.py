import tkinter as tk
from tkinter import filedialog
import xlwings as xw


# Global variable to store the workbook reference
workbook_excel = None


# Open The workbook for the first time
def open_excel_workbook(file_path):
    global workbook_excel
    if workbook_excel is None:
        app = xw.App(visible=True)
        workbook_excel = app.books.open(file_path)


def select_file():
    file_path = filedialog.askopenfilename()
    if file_path:
        file_entry.delete(0, tk.END)
        file_entry.insert(0, file_path)


def is_sub_string(substr, str) -> bool:
    i = 0
    j = 0

    while i < len(substr) and j < len(str):
        if str[j] == substr[i]:
            i += 1
        j += 1

    return i == len(substr)


def update_sheet_name_listbox(event):
    open_excel_workbook(file_entry.get())

    # Convert to uppercase for case-insensitive comparison
    input_text = sheet_name_entry.get().upper()

    available_sheets = [
        sheet.name
        for sheet in workbook_excel.sheets
        if is_sub_string(input_text, sheet.name.upper())
    ]

    sheet_name_listbox.delete(0, tk.END)

    for sheet_name in available_sheets:
        sheet_name_listbox.insert(tk.END, sheet_name)


# Function to update the sheet name entry field with the selected item from the listbox
def select_sheet_name(event):
    selected_sheet_name = sheet_name_listbox.get(sheet_name_listbox.curselection())
    sheet_name_entry.delete(0, tk.END)
    sheet_name_entry.insert(0, selected_sheet_name)


def search_excel():
    try:
        sheet_name = sheet_name_entry.get()
        ref_column = ref_column_entry.get()
        reference_value = ref_entry.get()

        open_excel_workbook(file_entry.get())

        sheet = workbook_excel.sheets[sheet_name]

        sheet.activate()

        if reference_value:
            reference_value = int(reference_value)

            row_num = None
            for i in range(1, 50):
                ref_cell_value = sheet.range(ref_column + str(i)).value
                if ref_cell_value == reference_value:
                    row_num = i
                    break

            if row_num is not None:
                sheet.range("A" + str(row_num) + ":AZ" + str(row_num)).select()
                result_label.config(text="Cell selected in the Excel file.")
            else:
                result_label.config(text="Reference value not found.")
        else:
            result_label.config(
                text="Sheet activated. Enter a reference value to search."
            )
    except Exception as e:
        result_label.config(text=f"Error: {str(e)}")


root = tk.Tk()
root.title("Excel Assistant")


file_label = tk.Label(root, text="File Path:")
file_label.pack()
file_entry = tk.Entry(root)
file_entry.pack()


select_file_button = tk.Button(root, text="Select File", command=select_file)
select_file_button.pack()

ref_column_label = tk.Label(root, text="Reference Column (e.g., 'AA'):")
ref_column_label.pack()
ref_column_entry = tk.Entry(root)
ref_column_entry.pack()

sheet_name_label = tk.Label(root, text="Sheet Name:")
sheet_name_label.pack()
sheet_name_entry = tk.Entry(root)
sheet_name_entry.pack()
sheet_name_entry.bind("<KeyRelease>", update_sheet_name_listbox)


sheet_name_listbox = tk.Listbox(root, height=5)
sheet_name_listbox.pack()
sheet_name_listbox.bind("<ButtonRelease-1>", select_sheet_name)


ref_label = tk.Label(root, text="Reference (Optional):")
ref_label.pack()
ref_entry = tk.Entry(root)
ref_entry.pack()


search_button = tk.Button(root, text="Search", command=search_excel)
search_button.pack()


result_label = tk.Label(root, text="")
result_label.pack()

root.mainloop()
