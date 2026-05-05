import openpyxl

file_path = '/Users/nayangarg/Desktop/DigitcomWebsiteRenovation/Old_Codebase_renovated_v6.1/PerformaInvoiceWork/debug_DC090.xlsx'
try:
    wb = openpyxl.load_workbook(file_path, data_only=True)
    ws = wb['JMS']
    print("\n--- debug_DC090 JMS Row 1-20 ---")
    for r in range(1, 21):
        row_vals = [ws.cell(row=r, column=c).value for c in range(1, 15)]
        print(f"Row {r}: {row_vals}")
except Exception as e:
    print(f"Error: {e}")
