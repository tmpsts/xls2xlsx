# xls2xlsx_file.py - Matthew Denton
# Takes in two arguments: source_path and destination_path
# Converts .xls file to .xlsx file from source_path to destination_path

import os
import sys
import win32com.client as win32
    
def convert_xls_to_xlsx(s, d) -> None:
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    wb = excel.Workbooks.Open(s)

    # FileFormat=51 is for .xlsx extension
    wb.SaveAs(d, FileFormat=51)
    wb.Close()
    excel.Application.Quit()
    os.remove(s)

if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("Usage: python xls2xlsx_fily.py <source> <destination>")
        sys.exit(1)

    source_path = sys.argv[1]
    destination_path = sys.argv[2]

    convert_xls_to_xlsx(source_path, destination_path)
