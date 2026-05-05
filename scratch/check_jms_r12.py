import openpyxl

file_path = '/Users/nayangarg/Desktop/DigitcomWebsiteRenovation/Old_Codebase_renovated_v6.1/PerformaInvoiceWork/50/DIGITCOM_AIRFIBER_DC062_JDPR 10-SEP-25 12-SEP-25  19-SEP-25_A6+B6.xlsx'
try:
    wb = openpyxl.load_workbook(file_path, data_only=True)
    ws = wb['JMS']
    print("\n--- JMS Row 12 (Indices 1 to 30) ---")
    for col in range(1, 31):
        val = ws.cell(row=12, column=col).value
        print(f"Col {col}: {val}")
except Exception as e:
    print(f"Error: {e}")
