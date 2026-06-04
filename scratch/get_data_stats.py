import pandas as pd
from data_loader import DataFactory

# Cohort 1: A6
a6_master = "../FinaliseBillingFormat/BILLING TRACKER.xlsx"
a6_mindump = "../FinaliseBillingFormat/MIN TRACKER.xlsx"
factory_a6 = DataFactory(a6_master)
factory_a6.sync_from_master()
factory_a6.sync_from_mindump(a6_mindump)

# Cohort 2: A6+B6
a6b6_master = "../a6b6_essentials/a6b6 billing.xlsx"
a6b6_mindump = "../a6b6_essentials/min tracker a6 b6.xlsx"
factory_a6b6 = DataFactory(a6b6_master)
factory_a6b6.sync_from_master()
factory_a6b6.sync_from_mindump(a6b6_mindump)

print("\n=== DATA SETS SUMMARY ===")
print(f"A6 Cohort (BILLING TRACKER):")
print(f"  Total Sites: {len(factory_a6.sites)}")
a6_dcs = set(str(s.dc_no).strip().upper() for s in factory_a6.sites.values() if s.dc_no)
print(f"  Total unique DCs: {len(a6_dcs)}")
print(f"  DCs present: {sorted(list(a6_dcs))[:10]}... (showing first 10)")
total_a6_disp = sum(len(s.dispatches) for s in factory_a6.sites.values())
print(f"  Total dispatches loaded: {total_a6_disp}")

print(f"\nA6+B6 Cohort (a6b6 billing.xlsx):")
print(f"  Total Sites: {len(factory_a6b6.sites)}")
a6b6_dcs = set(str(s.dc_no).strip().upper() for s in factory_a6b6.sites.values() if s.dc_no)
print(f"  Total unique DCs: {len(a6b6_dcs)}")
print(f"  DCs present: {sorted(list(a6b6_dcs))}")
total_a6b6_disp = sum(len(s.dispatches) for s in factory_a6b6.sites.values())
print(f"  Total dispatches loaded: {total_a6b6_disp}")
