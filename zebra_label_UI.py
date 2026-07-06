import pandas as pd
import win32print
import tkinter as tk
from tkinter import filedialog, messagebox

# ==============================
# PRINTER NAME (exact match)
# ==============================
PRINTER_NAME = "ZDesigner ZT231-203dpi ZPL"

selected_file = ""

# ==============================
# UTILITIES
# ==============================

def format_label_value(value):
    if pd.isna(value):
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value)

# ==============================
# SELECT EXCEL FILE
# ==============================

def select_excel():
    global selected_file
    selected_file = filedialog.askopenfilename(
        title="Select Excel File",
        filetypes=[("Excel Files", "*.xlsx")]
    )
    if selected_file:
        file_label.config(text=selected_file)

# ==============================
# PRINT LABELS
# ==============================

def print_labels():
    if not selected_file:
        messagebox.showerror("Error", "Please select an Excel file first")
        return

    try:
        df = pd.read_excel(selected_file)

        printer = win32print.OpenPrinter(PRINTER_NAME)
        win32print.StartDocPrinter(printer, 1, ("ZT231 Labels", None, "RAW"))
        win32print.StartPagePrinter(printer)

        rows = list(df.iterrows())

        for i in range(0, len(rows), 2):
            _, row1 = rows[i]
            price1 = format_label_value(row1["Price"])
            item_no1 = format_label_value(row1["Item No."])
            name1 = format_label_value(row1["Name of Item"])

            if i + 1 < len(rows):
                _, row2 = rows[i + 1]
                price2 = format_label_value(row2["Price"])
                item_no2 = format_label_value(row2["Item No."])
                name2 = format_label_value(row2["Name of Item"])
            else:
                price2 = ""
                item_no2 = ""
                name2 = ""

            zpl = f"""
^XA
^PW400
^LL200

^CF0,28

^FO0,30^FB400,1,0,C^FD{price1} - {item_no1}^FS
^FO0,55^FB400,1,0,C^FD{name1}^FS

^FO20,100^GB360,0,2^FS

^FO0,130^FB400,1,0,C^FD{price2} - {item_no2}^FS
^FO0,155^FB400,1,0,C^FD{name2}^FS

^XZ
"""
            win32print.WritePrinter(printer, zpl.encode())

        win32print.EndPagePrinter(printer)
        win32print.EndDocPrinter(printer)
        win32print.ClosePrinter(printer)

        messagebox.showinfo("Success", "Labels printed successfully")

    except Exception as e:
        messagebox.showerror("Error", str(e))

# ==============================
# UI DESIGN
# ==============================

def main():
    global file_label

    root = tk.Tk()
    root.title("Zebra Label Printer")
    root.geometry("520x220")
    root.resizable(False, False)

    tk.Label(root, text="Zebra ZT231 Label Printing", font=("Arial", 14, "bold")).pack(pady=10)

    tk.Button(
        root,
        text="Select Excel File",
        width=25,
        command=select_excel
    ).pack(pady=5)

    file_label = tk.Label(root, text="No file selected", fg="blue", wraplength=480)
    file_label.pack(pady=5)

    tk.Button(
        root,
        text="PRINT LABELS",
        width=25,
        height=2,
        bg="green",
        fg="white",
        command=print_labels
    ).pack(pady=15)

    root.mainloop()

if __name__ == "__main__":
    main()
