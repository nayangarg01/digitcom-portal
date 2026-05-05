import openpyxl

file_path = '/Users/nayangarg/Desktop/DigitcomWebsiteRenovation/Old_Codebase_renovated_v6.1/PerformaInvoiceWork/48/DIGITCOM_ AIRFIBER_DC083_ JDPR_31-OCT-25.xlsx'
try:
    wb = openpyxl.load_workbook(file_path, data_only=True)
    ws = wb['JMS']
    print(f"JMS H4: {ws['H4'].value}")
    
    ws_wcc = wb['Main WCC']
    print(f"Main WCC D29: {ws_wcc['D29'].value}")
except Exception as e:
    print(f"Error: {e}")
