import pandas as pd
from data_loader import DataFactory

factory = DataFactory("../a6b6_essentials/a6b6 billing.xlsx")
factory.sync_from_master()

# Filter sites for DC0128
dc_sites = [s for s in factory.sites.values() if str(s.dc_no).strip().upper() == 'DC0128']

print(f"DC0128 has {len(dc_sites)} sites:")
for s in dc_sites:
    print(f"  Site ID: {s.site_id}, Hop ID: {s.hop_id}, PMP ID: {s.pmp_id}")

# Load B6 dump directly to see what hop ends we have
print("\n--- Hop Ends for B6 matching ---")
for site_obj in dc_sites:
    hop_id = getattr(site_obj, 'hop_id', '')
    if hop_id and hop_id != 'N/A' and str(hop_id).upper() != 'NONE':
        clean_hop = str(hop_id).replace('_A6', '')
        if '-I-RJ-' in clean_hop:
            parts = clean_hop.split('-I-RJ-')
            ends = [parts[0], 'I-RJ-' + parts[1]]
        else:
            ends = [clean_hop]
        print(f"Site {site_obj.site_id} Hop Ends: {ends}")

print("\n--- Syncing MIN Dump ---")
# Let's sync from the specific min tracker in a6b6_essentials/
factory.sync_from_mindump("../a6b6_essentials/min tracker a6 b6.xlsx")

print("\n--- DC0128 B6 Dispatches matched ---")
for s in dc_sites:
    b6_dispatches = [d for d in s.dispatches if d['activity'] == 'B6']
    print(f"Site {s.site_id} (Hop: {s.hop_id}) has {len(b6_dispatches)} B6 dispatches:")
    pmp_set = set(d['pmp_id'] for d in b6_dispatches)
    print(f"  Matched B6 PMPs: {pmp_set}")
    for d in b6_dispatches[:5]:
        print(f"    PMP: {d['pmp_id']}, SAP: {d['sap_code']}, Qty: {d['quantity']}, Remarks: {d['remarks']}")
