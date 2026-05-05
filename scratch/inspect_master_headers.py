import openpyxl

def inspect_master():
    master_path = '/Users/nayangarg/Desktop/DigitcomWebsiteRenovation/Old_Codebase_renovated_v6.1/FinaliseBillingFormat/MASTER TRACKER FOR BILLING-AIRFIBER-RJST (1).xlsx'
    wb = openpyxl.load_workbook(master_path, data_only=True)
    ws = wb.active # Assuming first sheet
    print(f"Sheet: {ws.title}")
    
    # Search for DC number and WO column
    # Let's check headers in first 5 rows
    for r in range(1, 6):
        row = [str(c.value) for c in ws[r]]
        print(f"Row {r}: {row[:20]}") # Print first 20 columns

if __name__ == "__main__":
    inspect_master()
