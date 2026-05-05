import pandas as pd

mindump_path = '/Users/nayangarg/Desktop/DigitcomWebsiteRenovation/Old_Codebase_renovated_v6.1/InputBilling/MIN DUMP-RJST TILL 31 MAR 26.xlsx'
try:
    df_mindump = pd.read_excel(mindump_path, sheet_name='B6 DUMP')
    print("Columns in B6 DUMP (Pandas):")
    print(df_mindump.columns.tolist())
    print("\nSample rows:")
    print(df_mindump.head(3))
except Exception as e:
    print(f"Error: {e}")
