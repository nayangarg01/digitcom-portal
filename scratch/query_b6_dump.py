import pandas as pd

df = pd.read_excel("../a6b6_essentials/min tracker a6 b6.xlsx", sheet_name="b6")
print(f"Total B6 dump rows: {len(df)}")

ends = ['I-RJ-DDWN-ENB-I016', 'I-RJ-DDWN-ENB-I022']
for end in ends:
    mask = (
        df['ENB ID'].astype(str).str.contains(end, na=False) |
        df['Site ID'].astype(str).str.contains(end, na=False) |
        df['DWG'].astype(str).str.contains(end, na=False)
    )
    matching = df[mask]
    print(f"\nMatches for {end}: {len(matching)} rows")
    if not matching.empty:
        # Print unique COMMON ID and Site ID
        c_ids = matching['COMMON ID'].unique()
        s_ids = matching['Site ID'].unique()
        print(f"  Unique COMMON ID: {c_ids}")
        print(f"  Unique Site ID: {s_ids}")
        # Print a few rows
        print(matching[['ENB ID', 'Site ID', 'DWG', 'COMMON ID', 'SAP Code', 'No. Of Qty']].head(5))
