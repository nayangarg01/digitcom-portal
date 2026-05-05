import openpyxl

file_path = '/Users/nayangarg/Desktop/DigitcomWebsiteRenovation/Old_Codebase_renovated_v6.1/PerformaInvoiceWork/48/MATERIAL_REPORT_FOR_PERFORMA_INVOICE_NO_048_DIGITCOM.xlsx'
try:
    wb = openpyxl.load_workbook(file_path, data_only=True)
    ws = wb['1']
    print(f"Ref Row 6: {list(ws.iter_rows(min_row=6, max_row=6, values_only=True))[0]}")
except Exception as e:
    print(f"Error: {e}")
