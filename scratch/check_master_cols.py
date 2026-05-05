import pandas as pd
import os

master_path = '/Users/nayangarg/Desktop/DigitcomWebsiteRenovation/Old_Codebase_renovated_v6.1/InputBilling/MASTER TRACKER FOR BILLING-AIRFIBER-RJST (1).xlsx'
try:
    df = pd.read_excel(master_path, sheet_name='MASTER DPR')
    print(f"Columns: {df.columns.tolist()[:20]}")
except Exception as e:
    print(f"Error: {e}")
