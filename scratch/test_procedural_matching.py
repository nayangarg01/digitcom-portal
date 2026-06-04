import pandas as pd

df_sites = pd.read_excel("../a6b6_essentials/a6b6 billing.xlsx")
# Filter for DC0128
df_sites = df_sites[df_sites['BILLING FILE'].astype(str).str.strip().str.upper() == 'DC0128']
print(f"df_sites length: {len(df_sites)}")

df_mindump = pd.read_excel("../a6b6_essentials/min tracker a6 b6.xlsx", sheet_name="b6")

billing_pmp_ids = df_sites['PMP ID'].astype(str).str.strip().tolist()
print(f"billing_pmp_ids: {billing_pmp_ids}")

ann_pmp_ids = []

for _, site_row in df_sites.iterrows():
    hop_id = str(site_row.get('FB-FT HOP ID', '')).strip()
    print(f"\nProcessing Hop ID: {hop_id}")
    if not hop_id or hop_id.upper() == 'NONE': continue
    
    clean_hop = hop_id.replace('_A6', '')
    ends = []
    if '-I-RJ-' in clean_hop:
        parts = clean_hop.split('-I-RJ-')
        ends = [parts[0], 'I-RJ-' + parts[1]]
    else:
        ends = [clean_hop]
    print(f"  Ends: {ends}")
    
    matched_pmp_for_this_site = []
    for end in ends:
        if not end: continue
        mask = (
            df_mindump['ENB ID'].astype(str).str.contains(end, na=False) |
            df_mindump['Site ID'].astype(str).str.contains(end, na=False) |
            df_mindump['DWG'].astype(str).str.contains(end, na=False)
        )
        matched_rows = df_mindump[mask]
        print(f"    End: {end} matched {len(matched_rows)} rows in dump")
        for _, m_row in matched_rows.iterrows():
            pmp = str(m_row.get('COMMON ID', '')).strip()
            if not pmp or pmp.lower() in ['nan', 'none']:
                pmp = str(m_row.get('Site ID', '')).strip()
            
            if pmp and pmp.lower() not in ['nan', 'none']:
                if pmp not in matched_pmp_for_this_site:
                    matched_pmp_for_this_site.append(pmp)
    
    print(f"  Matched PMPs for this site: {matched_pmp_for_this_site}")
    for p in matched_pmp_for_this_site:
        if p not in ann_pmp_ids:
            ann_pmp_ids.append(p)

print(f"\nResulting ann_pmp_ids: {ann_pmp_ids}")
