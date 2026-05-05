import pandas as pd

mindump_path = '/Users/nayangarg/Desktop/DigitcomWebsiteRenovation/Old_Codebase_renovated_v6.1/InputBilling/MIN DUMP-RJST TILL 31 MAR 26.xlsx'
try:
    df_dump = pd.read_excel(mindump_path, sheet_name='A6 DUMP')
    target_site = 'I-RJ-NAWW-PMP-0018'
    row = df_dump[df_dump['Site ID'] == target_site]
    if not row.empty:
        print(f"Site: {target_site}")
        print(f"WBS ID: {row['WBS ID'].values[0]}")
        print(f"DWG: {row['DWG'].values[0]}")
    else:
        print(f"Site {target_site} not found in MINDUMP.")
except Exception as e:
    print(f"Error: {e}")
