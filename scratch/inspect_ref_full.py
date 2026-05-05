import openpyxl

def inspect_ref_full():
    ref_path = '/Users/nayangarg/Desktop/DigitcomWebsiteRenovation/Old_Codebase_renovated_v6.1/FinaliseBillingFormat/DIGITCOM_ AIRFIBER_DC0105_ JDPR_29-JAN-26_A6 (REJECT) 2.xlsx'
    wb_ref = openpyxl.load_workbook(ref_path, data_only=True)
    ws = wb_ref['Main WCC']
    
    print("--- Main WCC Content (Rows 1-40, Cols A-I) ---")
    for r in range(1, 41):
        row_vals = []
        for c in range(1, 11):
            val = ws.cell(row=r, column=c).value
            row_vals.append(f"{val}" if val is not None else "")
        print(f"Row {r:2}: {' | '.join(row_vals)}")

if __name__ == "__main__":
    inspect_ref_full()
