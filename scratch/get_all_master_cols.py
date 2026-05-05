import openpyxl

def inspect_all_headers():
    master_path = '/Users/nayangarg/Desktop/DigitcomWebsiteRenovation/Old_Codebase_renovated_v6.1/FinaliseBillingFormat/MASTER TRACKER FOR BILLING-AIRFIBER-RJST (1).xlsx'
    wb = openpyxl.load_workbook(master_path, data_only=True)
    ws = wb.active
    
    print("--- All Headers in Row 2 ---")
    row2 = [c.value for c in ws[2]]
    for i, val in enumerate(row2):
        if val:
            print(f"Col {i}: {val}")

if __name__ == "__main__":
    inspect_all_headers()
