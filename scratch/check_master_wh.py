import openpyxl

file_path = '/Users/nayangarg/Desktop/DigitcomWebsiteRenovation/Old_Codebase_renovated_v6.1/InputBilling/MASTER TRACKER FOR BILLING-AIRFIBER-RJST (1).xlsx'
try:
    wb = openpyxl.load_workbook(file_path, read_only=True, data_only=True)
    ws = wb['A6+B6 Billings']
    headers = [str(h).strip() for h in next(ws.iter_rows(max_row=1, values_only=True))]
    print(f"Headers in A6+B6 Billings: {headers}")
except Exception as e:
    print(f"Error: {e}")
