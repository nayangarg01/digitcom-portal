import pandas as pd
import sys
import os

sys.path.append("../Backend_Portal/scripts")
from generate_clean_billing import load_master_data, get_warehouse_name

master_path = "../a6b6_essentials/a6b6 billing.xlsx"
mindump_path = "../a6b6_essentials/min tracker a6 b6.xlsx"
dc_number = "DC0128"
activity = "A6_B6"
sub_activity = "B6"

df_sites, code_to_col_idx = load_master_data(master_path, dc_number, activity=activity)
print(f"df_sites columns: {df_sites.columns.tolist()}")

billing_pmp_ids = df_sites['PMP ID'].astype(str).str.strip().tolist()
print(f"billing_pmp_ids: {billing_pmp_ids}")

df_mindump = pd.read_excel(mindump_path, sheet_name="b6")
print(f"df_mindump columns: {df_mindump.columns.tolist()}")

ann_pmp_ids = []
for _, site_row in df_sites.iterrows():
    hop_id = str(site_row.get('FB-FT HOP ID', '')).strip()
    print(f"hop_id: {hop_id}")
    if not hop_id or hop_id.upper() == 'NONE': continue
    clean_hop = hop_id.replace('_A6', '')
    ends = []
    if '-I-RJ-' in clean_hop:
        parts = clean_hop.split('-I-RJ-')
        ends = [parts[0], 'I-RJ-' + parts[1]]
    else:
        ends = [clean_hop]
    print(f"ends: {ends}")
    matched_pmp_for_this_site = []
    for end in ends:
        if not end: continue
        mask = (
            df_mindump['ENB ID'].astype(str).str.contains(end, na=False) |
            df_mindump['Site ID'].astype(str).str.contains(end, na=False) |
            df_mindump['DWG'].astype(str).str.contains(end, na=False)
        )
        matched_rows = df_mindump[mask]
        print(f"  end {end} matched {len(matched_rows)} rows")
        for _, m_row in matched_rows.iterrows():
            pmp = str(m_row.get('COMMON ID', '')).strip()
            if not pmp or pmp.lower() in ['nan', 'none']:
                pmp = str(m_row.get('Site ID', '')).strip()
            if pmp and pmp.lower() not in ['nan', 'none']:
                if pmp not in matched_pmp_for_this_site:
                    matched_pmp_for_this_site.append(pmp)
    print(f"  matched_pmp_for_this_site: {matched_pmp_for_this_site}")
    for p in matched_pmp_for_this_site:
        if p not in ann_pmp_ids:
            ann_pmp_ids.append(p)

print(f"ann_pmp_ids before pivot: {ann_pmp_ids}")

# Match Site Logic for Pivot
def match_site(row):
    if sub_activity == 'B6':
        cid = str(row.get('COMMON ID', '')).strip()
        if not cid or cid.lower() in ['nan', 'none']:
            cid = str(row.get('Site ID', '')).strip()
        if cid in ann_pmp_ids:
            return cid
    else:
        wbs = str(row.get('WBS ID', ''))
        site_id = str(row.get('Site ID', ''))
        for pid in ann_pmp_ids:
            if pid in wbs or pid in site_id:
                return pid
    return None

df_mindump['Matched_PMP'] = df_mindump.apply(match_site, axis=1)
df_filtered = df_mindump[df_mindump['Matched_PMP'].notna()]
print(f"df_filtered rows: {len(df_filtered)}")

if not df_filtered.empty:
    pt = pd.pivot_table(df_filtered, values='No. Of Qty', index=['SAP Code', 'Material Description'], columns='Matched_PMP', aggfunc='sum', fill_value=0)
    print(f"pt columns: {pt.columns.tolist()}")
