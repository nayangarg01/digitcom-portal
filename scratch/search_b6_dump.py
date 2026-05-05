import openpyxl

file_path = '/Users/nayangarg/Desktop/DigitcomWebsiteRenovation/Old_Codebase_renovated_v6.1/InputBilling/MIN DUMP-RJST TILL 31 MAR 26.xlsx'
try:
    wb = openpyxl.load_workbook(file_path, read_only=True, data_only=True)
    ws = wb['B6 DUMP']
    
    count = 0
    for row in ws.iter_rows(min_row=2, values_only=True):
        row_str = " ".join([str(x) for x in row if x is not None])
        if 'I-RJ-RSNG-ENB-V002' in row_str or 'I-RJ-RSNG-ENB-B001' in row_str:
            print(row)
            count += 1
            if count > 5: break
    
    if count == 0:
        print("No matches found for DC0111 sites in B6 DUMP.")
            
except Exception as e:
    print(f"Error: {e}")
