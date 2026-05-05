import openpyxl

file_path = '/Users/nayangarg/Desktop/DigitcomWebsiteRenovation/Old_Codebase_renovated_v6.1/PerformaInvoiceWork/debug_DC0105_MainWCC_v2.xlsx'
try:
    wb = openpyxl.load_workbook(file_path, data_only=True)
    ws = wb['Main WCC']
    print("\n--- debug_DC0105_MainWCC_v2 JMS Row 1-40 ---")
    for r in range(1, 41):
        row_vals = [ws.cell(row=r, column=c).value for c in range(1, 10)]
        print(f"Row {r}: {row_vals}")
except Exception as e:
    print(f"Error: {e}")
