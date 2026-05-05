import openpyxl

file_path = '/Users/nayangarg/Desktop/DigitcomWebsiteRenovation/Old_Codebase_renovated_v6.1/InputBilling/MIN DUMP-RJST TILL 31 MAR 26.xlsx'
try:
    wb = openpyxl.load_workbook(file_path, read_only=True, data_only=True)
    ws = wb['B6 DUMP']
    
    print("\nNon-None COMMON IDs in B6 DUMP:")
    count = 0
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[0] is not None:
            print(row[0], row[5])
            count += 1
            if count > 10: break
    if count == 0:
        print("All COMMON IDs are None in B6 DUMP samples.")
            
except Exception as e:
    print(f"Error: {e}")
