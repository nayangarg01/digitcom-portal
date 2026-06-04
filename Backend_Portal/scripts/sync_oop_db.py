import sys
import os
import argparse

# Add BillingEngine_OOP to python search path
script_dir = os.path.dirname(os.path.abspath(__file__))
oop_dir = os.path.abspath(os.path.join(script_dir, "..", "..", "BillingEngine_OOP"))
sys.path.append(oop_dir)

try:
    from data_loader import DataFactory
except ImportError as e:
    print(f"CRITICAL: Failed to import BillingEngine_OOP. Ensure the BillingEngine_OOP directory exists. Error: {e}")
    sys.exit(1)

def main():
    print("=== OOP DATABASE SYNC STARTED ===")
    parser = argparse.ArgumentParser()
    parser.add_argument("master_path", help="Path to Master DPR Excel tracker")
    parser.add_argument("--mindump", default=None, help="Path to MIN dump Excel tracker")
    args = parser.parse_args()

    master_path = args.master_path
    mindump_path = args.mindump

    print(f"LOG: Master Tracker Path: {master_path}")
    if not os.path.exists(master_path):
        print(f"ERROR: Master Tracker file not found: {master_path}")
        sys.exit(1)

    if mindump_path:
        print(f"LOG: MIN Dump Path: {mindump_path}")
        if not os.path.exists(mindump_path):
            print(f"ERROR: MIN Dump file not found: {mindump_path}")
            sys.exit(1)

    print("LOG: Initializing DataFactory and loading database...")
    try:
        factory = DataFactory(master_path)
    except Exception as e:
        print(f"ERROR: Failed to load OOP Database: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)

    initial_count = len(factory.sites)
    print(f"LOG: Loaded existing database with {initial_count} sites.")

    print("LOG: Starting Master Tracker sync...")
    try:
        factory.sync_from_master()
    except Exception as e:
        print(f"ERROR: Master Tracker sync failed: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)

    post_master_count = len(factory.sites)
    added_sites = post_master_count - initial_count
    print(f"LOG: Master Tracker sync completed. Sites in database: {post_master_count} (Added: {added_sites})")

    if mindump_path:
        print("LOG: Starting MIN Dump sync...")
        try:
            factory.sync_from_mindump(mindump_path)
        except Exception as e:
            print(f"ERROR: MIN Dump sync failed: {e}")
            import traceback
            traceback.print_exc()
            sys.exit(1)
        print("LOG: MIN Dump sync completed.")
    else:
        print("LOG: Skipping MIN Dump sync (no file provided).")

    print("=== OOP DATABASE SYNC COMPLETED SUCCESSFULLY ===")
    print(f"SUMMARY: Total sites registered: {len(factory.sites)} (Added: {added_sites})")

if __name__ == "__main__":
    main()
