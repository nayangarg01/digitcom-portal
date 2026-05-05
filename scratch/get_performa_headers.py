import openpyxl

file_path = '/Users/nayangarg/Desktop/DigitcomWebsiteRenovation/Old_Codebase_renovated_v6.1/PerformaInvoiceWork/MATERIAL_REPORT_FOR_PERFORMA_INVOICE_NO_048_DIGITCOM.xlsx'
try:
    wb = openpyxl.load_workbook(file_path, read_only=True, data_only=True)
    ws = wb['1']
    print("\n--- Sheet '1' Headers (Row 1) ---")
    headers = [cell.value for cell in ws[1]]
    print(headers)
    
    ws_summary = wb['Summary sheet1']
    print("\n--- Sheet 'Summary sheet1' Headers (Row 1) ---")
    summary_headers = [cell.value for cell in ws_summary[1]]
    print(summary_headers)
except Exception as e:
    print(f"Error: {e}")
