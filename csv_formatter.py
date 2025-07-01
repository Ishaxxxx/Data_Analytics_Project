import pandas as pd
import tkinter as tk
from tkinter import filedialog
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Border, Side, Alignment
from win32com.client import Dispatch
from datetime import datetime
import os
from colorama import init, Fore, Style

init(autoreset=True)

# ============================
# Custom Logger
# ============================
def log(message, level="info"):
    if level == "info":
        print(Fore.CYAN + "[INFO] " + Style.RESET_ALL + message)
    elif level == "success":
        print(Fore.GREEN + "[SUCCESS] " + Style.RESET_ALL + message)
    elif level == "error":
        print(Fore.RED + "[ERROR] " + Style.RESET_ALL + message)
    elif level == "step":
        print(Fore.YELLOW + "[STEP] " + Style.RESET_ALL + message)

# ============================
# Custom Info Dialog
# ============================
def show_custom_info(title, message):
    window = tk.Tk()
    window.title(title)
    window.configure(bg="#4CAF50")
    window.geometry("450x250")
    window.resizable(False, False)

    label = tk.Label(window, text=message, bg="#4CAF50", fg="white",
                     font=("Helvetica", 14), wraplength=400, justify="center")
    label.pack(expand=True, padx=20, pady=20)

    button = tk.Button(window, text="Awesome! 🚀", command=window.destroy,
                       bg="white", fg="black", font=("Helvetica", 12), width=15)
    button.pack(pady=15)

    window.mainloop()

# ============================
# File and Folder Selectors
# ============================
def select_csv_file():
    show_custom_info("📂 File Selection", "Please select your CSV file to continue.")
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(title="Select CSV File", filetypes=[("CSV Files", "*.csv")])
    root.destroy()
    return file_path

def select_output_folder():
    show_custom_info("📀 Output Folder", "Please select the folder where you want to save the Excel file.")
    root = tk.Tk()
    root.withdraw()
    folder_path = filedialog.askdirectory(title="Select Output Folder")
    root.destroy()
    return folder_path

# ============================
# Progress Popup (No threading)
# ============================
def show_progress():
    progress = tk.Tk()
    progress.title("⏳ Processing...")
    progress.geometry("400x100")
    progress.configure(bg="white")
    progress.resizable(False, False)

    label = tk.Label(progress, text="Processing your file... Please wait.", font=("Helvetica", 12), bg="white")
    label.pack(pady=20)

    progress.after(3000, progress.destroy)
    progress.mainloop()

# ============================
# CSV to Excel Processing
# ============================
def process_csv(csv_path, output_path):
    log("Reading CSV file...", "step")
    try:
        df = pd.read_csv(csv_path, encoding='utf-8', low_memory=False)
    except UnicodeDecodeError:
        log("UTF-8 encoding failed. Trying latin1...", "step")
        df = pd.read_csv(csv_path, encoding='latin1', low_memory=False)

    required_columns = [
        'Case Number', 'Created Date', 'Customer Name', 'Customer Phone', 'Street',
        'Zip/Postal Code', 'Customer Complaint', 'Product Description',
        'LineItem Status', 'Technician Name', 'WO Status'
    ]

    df = df[required_columns]
    df = df[df['WO Status'] == 'New']

    log("Calculating SLA...", "step")
    df['Created Date'] = pd.to_datetime(df['Created Date'], dayfirst=True, errors='coerce')
    df['SLA'] = (datetime.now() - df['Created Date']).dt.days.fillna(0).astype(int)

    df = df[['Case Number', 'SLA', 'Customer Name', 'Customer Phone', 'Street',
             'Zip/Postal Code', 'Customer Complaint', 'Product Description',
             'LineItem Status', 'Technician Name']]
    df['Remarks'] = ''
    df['Count'] = 1  # Add column for counting in pivot

    log("Saving to Excel with styling...", "step")
    save_to_excel(df, output_path)

# ============================
# Excel Styling
# ============================
def save_to_excel(df, excel_path):
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"

    ws.append(df.columns.tolist())
    for row in df.values.tolist():
        ws.append(row)

    header_fill = PatternFill(start_color="4CAF50", end_color="4CAF50", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF")
    border = Border(left=Side(style='thin'), right=Side(style='thin'),
                    top=Side(style='thin'), bottom=Side(style='thin'))
    align = Alignment(horizontal='center', vertical='center', wrap_text=True)

    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.border = border
        cell.alignment = align

    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.border = border
            cell.alignment = align

    for row in ws.iter_rows(min_row=2, min_col=2, max_col=2):
        for cell in row:
            if isinstance(cell.value, int) and cell.value >= 1:
                cell.fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")

    ws.auto_filter.ref = f"A1:L{ws.max_row}"

    widths = [12, 8, 20, 15, 50, 15, 20, 30, 15, 20, 20, 10]
    for i, width in enumerate(widths, 1):
        ws.column_dimensions[chr(64 + i)].width = width

    wb.save(excel_path)
    log("Excel file saved with styling.", "success")

# ============================
# Excel COM: Sort, Filter, Pivot
# ============================
def excel_sort_filter_pivot(excel_path):
    log("Opening Excel to apply sorting and pivot table...", "step")
    excel = Dispatch('Excel.Application')
    excel.Visible = True
    excel.DisplayAlerts = False

    wb = excel.Workbooks.Open(os.path.abspath(excel_path))
    ws = wb.Sheets("Sheet1")
    last_row = ws.UsedRange.Rows.Count
    data_range = ws.Range(f"A1:L{last_row}")

    if ws.AutoFilterMode:
        ws.AutoFilterMode = False

    data_range.AutoFilter()
    data_range.Sort(Key1=ws.Range("B1"), Order1=2, Header=1)
    ws.Range("A1").AutoFilter(Field=9, Criteria1="New")

    ws_pivot = wb.Sheets.Add()
    ws_pivot.Name = "Pivot_View"

    pivot_cache = wb.PivotCaches().Create(1, data_range)
    pivot_table = pivot_cache.CreatePivotTable(ws_pivot.Range("B4"), "SLA_Pivot")

    pivot_table.PivotFields("Technician Name").Orientation = 1
    pivot_table.PivotFields("SLA").Orientation = 2
    pivot_table.PivotFields("LineItem Status").Orientation = 3

    try:
        pf = pivot_table.PivotFields("Count")
        pivot_table.AddDataField(pf, "Count of Cases", -4157)
    except Exception as e:
        log(f"Pivot AddDataField failed: {e}", "error")
        try:
            available_fields = [pivot_table.PivotFields(i + 1).Name for i in range(pivot_table.PivotFields().Count)]
            log(f"Available Pivot Fields: {available_fields}", "info")
        except Exception as sub_e:
            log(f"Couldn't retrieve pivot field names: {sub_e}", "error")

    try:
        excel.CommandBars("PivotTable").Visible = True
    except Exception:
        log("Could not enable PivotTable field list (UI-dependent).", "info")

    wb.Save()
    wb.Close(SaveChanges=True)
    excel.Quit()
    log("Pivot table created successfully with filters and field pane.", "success")

# ============================
# Main Application Flow
# ============================
def main():
    print(Fore.GREEN + "\n===============================")
    print(Fore.CYAN + "🚀 Welcome to CSV to Excel Formatter - Google Developer Style")
    print(Fore.CYAN + "✨ Built with love to make your work smarter!")
    print(Fore.GREEN + "===============================\n")

    csv_path = select_csv_file()
    if not csv_path:
        log("No CSV file selected. Exiting...", "error")
        return

    output_folder = select_output_folder()
    if not output_folder:
        log("No output folder selected. Exiting...", "error")
        return

    filename = f"Output_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    excel_path = os.path.join(output_folder, filename)

    show_progress()

    log("Starting CSV to Excel conversion...", "step")
    process_csv(csv_path, excel_path)
    excel_sort_filter_pivot(excel_path)

    os.startfile(excel_path)

    show_custom_info("✅ Success", f"Your Excel file has been successfully created and opened!\n\nFile Path:\n{excel_path}")

    log("All done! File opened successfully.", "success")
    print(Fore.GREEN + "\n===============================")
    print(Fore.CYAN + "🎉 Process Completed Successfully!")
    print(Fore.CYAN + "🚀 Thank you for using the CSV Formatter!")
    print(Fore.GREEN + "===============================\n")

if __name__ == "__main__":
    main()