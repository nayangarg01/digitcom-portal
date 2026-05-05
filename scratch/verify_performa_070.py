import openpyxl

file_path = '/Users/nayangarg/Desktop/DigitcomWebsiteRenovation/Old_Codebase_renovated_v6.1/PerformaInvoiceWork/70/MATERIAL_REPORT_FOR_PERFORMA_INVOICE_NO_070_DIGITCOM.xlsx'
try:
    wb = openpyxl.load_workbook(file_path, data_only=True)
    ws_sum = wb['Summary sheet1']
    print("\n--- Performa 070 Summary Sheet ---")
    for row in ws_sum.iter_rows(max_row=10, values_only=True):
        print(row)
except Exception as e:
    print(f"Error: {e}")
