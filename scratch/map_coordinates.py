import openpyxl

ref_file = '/Users/nayangarg/Desktop/DigitcomWebsiteRenovation/Old_Codebase_renovated_v6.1/FinaliseBillingFormat/DIGITCOM_ AIRFIBER_DC0105_ JDPR_29-JAN-26_A6 (REJECT) 2.xlsx'

try:
    wb = openpyxl.load_workbook(ref_file, data_only=True)
    ws = wb['Main WCC']
    print("\n--- Detailed Coordinate Mapping of Reference Main WCC ---")
    for r in range(1, 40):
        row_vals = []
        for c in range(1, 10):
            cell = ws.cell(row=r, column=c)
            val = cell.value
            row_vals.append(f"{val}")
        print(f"Row {r}: {row_vals}")
except Exception as e:
    print(f"Error: {e}")
