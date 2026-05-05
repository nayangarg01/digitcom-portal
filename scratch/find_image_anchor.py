import openpyxl

ref_file = '/Users/nayangarg/Desktop/DigitcomWebsiteRenovation/Old_Codebase_renovated_v6.1/FinaliseBillingFormat/DIGITCOM_ AIRFIBER_DC0105_ JDPR_29-JAN-26_A6 (REJECT) 2.xlsx'

try:
    from openpyxl import load_workbook
    wb = load_workbook(ref_file)
    ws = wb['Main WCC']
    print(f"Number of images in Main WCC: {len(ws._images)}")
    for img in ws._images:
        print(f"Image anchor: {img.anchor}")
except Exception as e:
    print(f"Error: {e}")
