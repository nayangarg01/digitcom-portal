import os
import pickle
import sys

from data_loader import DataFactory

def test_persistence():
    db_path = "sites_db.pkl"
    if os.path.exists(db_path):
        os.remove(db_path)
        print(f"Removed existing {db_path}")

    # Use the essentials folders files
    a6_master = "../a6b6_essentials/a6 billing.xlsx"
    a6b6_master = "../a6b6_essentials/a6b6 billing.xlsx"
    mindump = "../a6b6_essentials/min tracker a6 b6.xlsx"

    print("\n--- PHASE 1: Load and sync A6 Master ---")
    factory = DataFactory(a6_master, db_path=db_path)
    factory.sync_from_master()
    
    a6_site_count = len(factory.sites)
    print(f"Sync complete. Site count in factory: {a6_site_count}")
    
    assert os.path.exists(db_path), "Database pickle file was not created!"
    print("Database pickle file successfully created.")

    # Load file and inspect
    with open(db_path, "rb") as f:
        loaded_sites = pickle.load(f)
    print(f"Successfully verified pickle file contains {len(loaded_sites)} sites.")
    assert len(loaded_sites) == a6_site_count, "Pickle contents size mismatch!"

    print("\n--- PHASE 2: Re-initialize and sync from A6+B6 Master (Incremental Merging) ---")
    # This should load A6 sites, then add/merge A6+B6 sites
    factory2 = DataFactory(a6b6_master, db_path=db_path)
    print(f"Before sync, loaded site count: {len(factory2.sites)}")
    assert len(factory2.sites) == a6_site_count, "Should have loaded pre-existing A6 sites"
    
    factory2.sync_from_master()
    merged_site_count = len(factory2.sites)
    print(f"After syncing A6+B6 master, merged site count: {merged_site_count}")
    assert merged_site_count > a6_site_count, "Site count should have increased!"

    print("\n--- PHASE 3: Sync MIN Dump (Material Dispatches) ---")
    # Sync MIN Dump and verify it saves dispatches
    factory2.sync_from_mindump(mindump)
    
    # Verify some site now has dispatches
    sites_with_disp = [s for s in factory2.sites.values() if len(s.dispatches) > 0]
    print(f"Sites with dispatches: {len(sites_with_disp)}")
    assert len(sites_with_disp) > 0, "No sites received dispatches!"
    
    # Reload from database to verify dispatches persisted in pickle
    print("\n--- PHASE 4: Verify Pickle Persistence of Dispatches ---")
    factory3 = DataFactory(a6b6_master, db_path=db_path)
    sites_with_disp_reloaded = [s for s in factory3.sites.values() if len(s.dispatches) > 0]
    print(f"Reloaded: Sites with dispatches: {len(sites_with_disp_reloaded)}")
    assert len(sites_with_disp_reloaded) == len(sites_with_disp), "Dispatches did not persist in pickle database!"

    print("\nSUCCESS: All persistence and merging checks passed!")

if __name__ == "__main__":
    test_persistence()
