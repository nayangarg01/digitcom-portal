import pandas as pd
import sys

file_path = '/Users/nayangarg/Desktop/DigitcomWebsiteRenovation/Old_Codebase_renovated_v6.1/InputBilling/MASTER TRACKER FOR BILLING-AIRFIBER-RJST (1).xlsx'
try:
    xl = pd.ExcelFile(file_path)
    print(f"Sheets: {xl.sheet_names}")
    df = pd.read_excel(file_path, sheet_name='A6+B6 Billings', header=None)
    print("\nFirst 10 rows of 'A6+B6 Billings':")
    print(df.head(10).to_string())
    
    # Check column names in row 1 (0-indexed)
    if len(df) > 1:
        print("\nHeaders (Row 1):")
        print(df.iloc[1].tolist())
except Exception as e:
    print(f"Error: {e}")
