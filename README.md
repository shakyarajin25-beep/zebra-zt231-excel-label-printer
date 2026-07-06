# Zebra Label Printer

Python GUI script to print labels on a Zebra ZT231 printer from an Excel file.

## Requirements

- Python 3
- pandas
- pywin32
- tkinter (typically included with Windows Python)

## Usage

1. Install dependencies:
   ```powershell
   pip install pandas pywin32
   ```
2. Run the script:
   ```powershell
   python zebra_label_UI.py
   ```
3. Choose an Excel file and click `PRINT LABELS`.

## Excel format

The Excel file should contain columns:
- `Price`
- `Item No.`
- `Name of Item`
