import openpyxl
import os

def inspect_files():
    master_path = '/Users/nayangarg/Desktop/DigitcomWebsiteRenovation/Old_Codebase_renovated_v6.1/FinaliseBillingFormat/MASTER TRACKER FOR BILLING-AIRFIBER-RJST (1).xlsx'
    ref_path = '/Users/nayangarg/Desktop/DigitcomWebsiteRenovation/Old_Codebase_renovated_v6.1/FinaliseBillingFormat/DIGITCOM_ AIRFIBER_DC0105_ JDPR_29-JAN-26_A6 (REJECT) 2.xlsx'
    
    print(f"--- Inspecting Master Tracker: {os.path.basename(master_path)} ---")
    try:
        wb_master = openpyxl.load_workbook(master_path, data_only=True)
        sheet = wb_master.active
        print(f"Active Sheet: {sheet.title}")
        # Print first 2 rows
        for row in sheet.iter_rows(min_row=1, max_row=2, values_only=True):
            print(row)
    except Exception as e:
        print(f"Error reading Master Tracker: {e}")

    print(f"\n--- Inspecting Reference File: {os.path.basename(ref_path)} ---")
    try:
        wb_ref = openpyxl.load_workbook(ref_path, data_only=True)
        print(f"Sheets: {wb_ref.sheetnames}")
        if 'Main WCC' in wb_ref.sheetnames:
            ws = wb_ref['Main WCC']
            print("Main WCC found. Inspecting key cells...")
            # I'll check common cells for Site Count, Completion Date, WO Number
            # In generate_clean_billing.py:
            # Site Count: C27:D27
            # Completion Date: F27:H27
            # WO Number: C25:D25
            
            cells_to_check = ['B23', 'C23', 'E23', 'F23', 'B25', 'C25', 'E25', 'F25', 'B27', 'C27', 'E27', 'F27']
            for coord in cells_to_check:
                print(f"{coord}: {ws[coord].value}")
        else:
            print("Main WCC NOT found in reference file.")
    except Exception as e:
        print(f"Error reading Reference File: {e}")

if __name__ == "__main__":
    inspect_files()
