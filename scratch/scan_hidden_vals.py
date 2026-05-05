import openpyxl

file_path = '/Users/nayangarg/Desktop/DigitcomWebsiteRenovation/Old_Codebase_renovated_v6.1/PerformaInvoiceWork/48/MATERIAL_REPORT_FOR_PERFORMA_INVOICE_NO_048_DIGITCOM.xlsx'
try:
    wb = openpyxl.load_workbook(file_path, data_only=True)
    ws = wb['1']
    # Scan columns 5, 6, 7 (E, F, G) for any non-null values
    for row in ws.iter_rows(min_row=2, max_row=100, values_only=True):
        if any(x is not None for x in row[4:7]):
            print(f"Row {row[0]}: {row[4:7]}")
except Exception as e:
    print(f"Error: {e}")
