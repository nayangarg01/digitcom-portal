import openpyxl
import os

folder = "../Backend_Portal/uploads"
for f in os.listdir(folder):
    if f.endswith(".xlsx"):
        path = os.path.join(folder, f)
        try:
            wb = openpyxl.load_workbook(path, read_only=True)
            print(f"File: {f}, Sheets: {wb.sheetnames}")
        except Exception as e:
            print(f"File: {f}, Error: {e}")
