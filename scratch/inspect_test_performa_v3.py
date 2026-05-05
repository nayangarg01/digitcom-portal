import openpyxl

file_path = '/Users/nayangarg/Desktop/DigitcomWebsiteRenovation/Old_Codebase_renovated_v6.1/PerformaInvoiceWork/Performa_Test_v3.xlsx'
try:
    wb = openpyxl.load_workbook(file_path, data_only=True)
    ws_sum = wb['Summary sheet1']
    print("\n--- Summary Sheet Sample Rows (Row 1-20) ---")
    for row in ws_sum.iter_rows(max_row=20, values_only=True):
        print(row)
except Exception as e:
    print(f"Error: {e}")
