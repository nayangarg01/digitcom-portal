import openpyxl

file_path = '/Users/nayangarg/Desktop/DigitcomWebsiteRenovation/Old_Codebase_renovated_v6.1/PerformaInvoiceWork/48/DIGITCOM_ AIRFIBER_DC083_ JDPR_31-OCT-25.xlsx'
try:
    wb = openpyxl.load_workbook(file_path, read_only=True, data_only=True)
    ws = wb['Main WCC']
    print("\n--- Main WCC Rows 1-30 ---")
    for row in ws.iter_rows(max_row=30, values_only=True):
        print(row)
except Exception as e:
    print(f"Error: {e}")
